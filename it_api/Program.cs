//using it.Data;
using Microsoft.EntityFrameworkCore;
using Microsoft.AspNetCore.Identity;
using Vue.Data;
using Vue.Models;
using Vue.Services;
using Microsoft.AspNetCore.Mvc.Infrastructure;
using Microsoft.AspNetCore.Authorization;
using Microsoft.Extensions.FileProviders;
using System.Net;
//using it.Services;
using System.Text.Json.Serialization;
using it_api.Services;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.IdentityModel.Tokens;
using System.Text;
using Microsoft.AspNetCore.Http.Features;
//using Vue.Middleware;

namespace Vue
{
    public class Program
    {
        public static string description { get; set; } = "";
        private static string MyAllowSpecificOrigins = "tran";
        public static void Main(string[] args)
        {
            Startup(args);

        }
        public static WebApplicationBuilder CreateDefaultBuilder(string[] args)
        {
            var builder = WebApplication.CreateBuilder(args);

            ConfigurationManager configuration = builder.Configuration;

            var connectionString = builder.Configuration.GetConnectionString("ItConnection") ?? throw new InvalidOperationException("Connection string 'ItConnection' not found.");
            var EsignConnectionString = builder.Configuration.GetConnectionString("EsignConnection") ?? throw new InvalidOperationException("Connection string 'EsignConnectionString' not found.");


            builder.Services.AddSingleton<IActionContextAccessor, ActionContextAccessor>();
            builder.Services.AddControllersWithViews().AddJsonOptions(x =>
              x.JsonSerializerOptions.ReferenceHandler = ReferenceHandler.IgnoreCycles);

            builder.Services.AddDbContext<IdentityContext>(options =>
                options.UseSqlServer(EsignConnectionString));
            builder.Services.AddDefaultIdentity<UserModel>(options => options.SignIn.RequireConfirmedAccount = false).AddRoles<IdentityRole>()
                .AddEntityFrameworkStores<IdentityContext>(); ;



            builder.Services.AddDbContext<ItContext>(options =>
              options.UseSqlServer(connectionString)
              );

            builder.Services.AddScoped<AesOperation, AesOperation>();
            builder.Services.AddScoped<ViewRender, ViewRender>();

            builder.Services.AddScoped<LoginMailPyme, LoginMailPyme>();

            builder.Services.AddScoped<AuthManager, AuthManager>();

            builder.Services.Configure<FormOptions>(x =>
            {
                // Password settings.
                x.BufferBody = false;
                x.KeyLengthLimit = 2048; // 2 KiB
                x.ValueLengthLimit = 4194304; // 32 MiB
                x.ValueCountLimit = 2048;// 1024
                x.MultipartHeadersCountLimit = 32; // 16
                x.MultipartHeadersLengthLimit = 32768; // 16384
                x.MultipartBoundaryLengthLimit = 256; // 128
                x.MultipartBodyLengthLimit = 134217728; // 128 MiB
            });

            //builder.Services.ConfigureApplicationCookie(options =>
            //{
            //    // Cookie settings
            //    options.Cookie.HttpOnly = false;
            //    options.ExpireTimeSpan = TimeSpan.FromHours(Int64.Parse(configuration["JWT:Expire"]));

            //    options.LoginPath = "/Identity/Account/Login";
            //    options.AccessDeniedPath = "/Identity/Account/AccessDenied";
            //    options.SlidingExpiration = true;
            //});
            builder.Services.AddAuthentication(options =>
            {
                options.DefaultAuthenticateScheme = JwtBearerDefaults.AuthenticationScheme;
                options.DefaultChallengeScheme = JwtBearerDefaults.AuthenticationScheme;
            })
            .AddJwtBearer(options =>
            {
                options.RequireHttpsMetadata = false;
                options.SaveToken = true;
                options.TokenValidationParameters = new TokenValidationParameters
                {
                    ValidateIssuer = true,
                    ValidateAudience = true,
                    ValidateLifetime = true,
                    ValidateIssuerSigningKey = true,
                    ValidIssuer = configuration["JWT:ValidIssuer"],
                    ValidAudience = configuration["JWT:ValidAudience"],
                    IssuerSigningKey = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(configuration["JWT:Secret"]))
                };
            });



            builder.Services.AddCors(options =>
            {
                options.AddPolicy(name: MyAllowSpecificOrigins,
                                  policy =>
                                  {
                                      policy.WithOrigins("*").AllowAnyMethod().AllowAnyHeader();
                                  });
            });
            return builder;
        }
        public static void Startup(string[] args)
        {
            var builder = CreateDefaultBuilder(args);
            var app = builder.Build();

            //app.UseMiddleware<CheckTokenMiddleware>();
            app.UseDeveloperExceptionPage();
            // Configure the HTTP request pipeline.
            app.UseStaticFiles();
            app.UseRouting();
            app.UseCors(MyAllowSpecificOrigins);

            app.UseAuthentication();
            app.UseAuthorization();

            //app.UseStaticFiles(new StaticFileOptions
            //{

            //    FileProvider = new PhysicalFileProvider(Path.Combine(builder.Environment.ContentRootPath, "frontend", "src")),
            //    RequestPath = "/src",
            //    OnPrepareResponse = ctx =>
            //    {
            //    }
            //});
            app.UseStaticFiles(new StaticFileOptions
            {
                FileProvider = new PhysicalFileProvider(builder.Configuration["Source:Path_Private"]),
                RequestPath = "/private",
                OnPrepareResponse = ctx =>
                {
                    var token = builder.Configuration["Key_Access"];
                    var token_query = ctx.Context.Request.Query["token"].ToString();

                    if (!ctx.Context.User.Identity.IsAuthenticated && token_query != token)
                    {
                        ctx.Context.Response.Redirect("/Identity/Account/Login");
                    }
                }
            });
            app.UseEndpoints(endpoints =>
            {
                endpoints.MapControllerRoute(
                   name: "areas",
                   pattern: "{area:exists}/{controller=Home}/{action=Index}/{id?}"
                 );
                endpoints.MapControllerRoute(
                    name: "default",
                    pattern: "{controller}/{action=Index}/{id?}");

            });
            app.Run();
        }
    }
}