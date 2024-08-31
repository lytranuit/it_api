



using Holdtime.Models;
using Info.Models;
using it_template.Areas.Trend.Controllers;
using iText.Commons.Actions.Contexts;
using iText.Commons.Utils;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using PH.WorkingDaysAndTimeUtility.Configuration;
using PH.WorkingDaysAndTimeUtility;
using System.Data;
using System.Dynamic;
using System.Text.Json.Serialization;
using Vue.Data;
using Vue.Models;
using Vue.Services;
using Spire.Xls;
using System.Security.Policy;
using Microsoft.EntityFrameworkCore.Infrastructure;
using System.Diagnostics;

namespace it_template.Areas.Info.Controllers
{

    [Authorize(Roles = "Administrator,HR")]
    public class SalaryController : BaseController
    {
        private readonly IConfiguration _configuration;
        private UserManager<UserModel> UserManager;
        private readonly TinhCong _tinhcong;
        private readonly ViewRender _view;
        private readonly ItContext _itContext;

        public SalaryController(NhansuContext context, AesOperation aes, TinhCong tinhcong, IConfiguration configuration, UserManager<UserModel> UserMgr, ViewRender view, ItContext itContext) : base(context, aes)
        {
            _configuration = configuration;
            UserManager = UserMgr;
            _tinhcong = tinhcong;
            _view = view;
            _itContext = itContext;
            var listener = _itContext.GetService<DiagnosticSource>();
            (listener as DiagnosticListener).SubscribeWithAdapter(new CommandInterceptor());
        }

        [HttpPost]
        public async Task<JsonResult> Delete(string id)
        {
            var Model = _context.SalaryModel.Where(d => d.id == id).FirstOrDefault();
            Model.deleted_at = DateTime.Now;
            _context.Update(Model);
            _context.SaveChanges();
            return Json(new { success = true, data = Model });
        }

        [HttpPost]
        public async Task<JsonResult> tinhluong(string id)
        {
            var SalaryUserModel = _context.SalaryUserModel.Where(d => d.salary_id == id).ToList();
            var Users = SalaryUserModel.Select(d => d.person_id).ToList();
            var Model = _context.SalaryModel.Where(d => d.id == id).FirstOrDefault();
            var date_from = Model.date_from.Value;
            var date_to = Model.date_to.Value;
            var data_post = _context.PersonnelModel.Where(d => Users.Contains(d.id)).ToList();

            var list_data_cong = _tinhcong.cong(data_post, date_from, date_to);
            var viewPath = "wwwroot/report/excel/Bảng lương.xlsx";
            var documentPath = "/private/info/data/" + DateTime.Now.ToFileTimeUtc() + ".xlsx";
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(viewPath);
            Worksheet sheet = workbook.Worksheets[0];
            var now = DateTime.Now;
            sheet.Range["O6"].Value = $"Tháng {date_to.ToString("MM")} Năm {date_to.ToString("yyyy")} (công tính từ ngày {date_from.ToString("dd/MM/yy")}-{date_to.ToString("dd/MM/yy")})";
            sheet.Range["AA3"].Value = $"Đông Hòa, ngày {now.ToString("dd")} tháng {now.ToString("MM")} năm {now.ToString("yyyy")}";
            int stt = 0;
            var start_r = 12;
            //DataTable dt = new DataTable();
            //dt.Columns.Add("stt", typeof(int));
            //dt.Columns.Add("manv", typeof(string));
            //dt.Columns.Add("hovaten", typeof(string));
            //dt.Columns.Add("chucvu", typeof(string));
            //dt.Columns.Add("phanloai", typeof(string));
            //dt.Columns.Add("ngaycongchuan", typeof(int));
            //dt.Columns.Add("luongcb", typeof(decimal));
            //dt.Columns.Add("luongdongbh", typeof(decimal));
            //dt.Columns.Add("tc_trangphuc", typeof(decimal));
            //dt.Columns.Add("tc_dienthoai", typeof(decimal));
            //dt.Columns.Add("tc_xangxe", typeof(decimal));
            //dt.Columns.Add("tc_antrua", typeof(decimal));
            //dt.Columns.Add("tc_chuyencan", typeof(decimal));
            //dt.Columns.Add("tc_anca", typeof(decimal));
            //dt.Columns.Add("tc_trachnhiem", typeof(decimal));
            //dt.Columns.Add("tongtrocap", typeof(decimal));
            //dt.Columns.Add("tien_luong_kpi", typeof(decimal));
            //dt.Columns.Add("congthucte", typeof(decimal));
            //dt.Columns.Add("bosung", typeof(decimal));
            //dt.Columns.Add("tongthunhap", typeof(decimal));

            //dt.Columns.Add("pc_trangphuc", typeof(decimal));
            //dt.Columns.Add("pc_anca", typeof(decimal));
            //dt.Columns.Add("thunhapchiuthue", typeof(decimal));

            //dt.Columns.Add("tncn_banthan", typeof(decimal));
            //dt.Columns.Add("tncn_songuoiphuthuoc", typeof(int));
            //dt.Columns.Add("tncn_nguoiphuthuoc", typeof(decimal));
            //dt.Columns.Add("tncn_bhxh", typeof(decimal));
            //dt.Columns.Add("thunhaptinhthue", typeof(decimal));
            //dt.Columns.Add("thue_tncn", typeof(decimal));
            //dt.Columns.Add("dpcd", typeof(decimal));
            //dt.Columns.Add("thuclanh", typeof(decimal));
            //dt.Columns.Add("tamungdot1", typeof(decimal));
            //dt.Columns.Add("conlai", typeof(decimal));
            sheet.InsertRow(start_r + 1, list_data_cong.Count(), InsertOptionsType.FormatAsAfter);
            foreach (var item in list_data_cong)
            {
                var record = data_post.Where(d => d.MANV == item["MANV"]).FirstOrDefault();
                var chucvu = _context.PositionModel.Where(d => d.MACHUCVU == record.MACHUCVU).FirstOrDefault();
                var bophan = _context.DepartmentModel.Where(d => d.MAPHONG == record.MAPHONG).FirstOrDefault();

                var congtrongthang = item["tongcong"];
                var congthucte = item["tong"];
                var tenchucvu = chucvu != null ? chucvu.TENCHUCVU : "";
                var phanloai = bophan != null ? bophan.TENPHONG : "";
                var manv = record.MANV;
                var hovaten = record.HOVATEN;
                var luongcb = record.tien_luong ?? 0;
                var pc_xangxe = record.pc_khac ?? 0;
                var pc_trachnhiem = record.pc_trachnhiem ?? 0;
                var luongkpi = record.tien_luong_kpi ?? 0;
                var TNCN_banthan = 11000000;
                var TNCN_nguoiphuthuoc = record.nguoiphuthuoc ?? 0;
                var tamungdot1 = record.tien_luong_dot1 ?? 0;


                ///
                //DataRow dr1 = dt.NewRow();
                //dr1["stt"] = (++stt);
                //dr1["manv"] = manv;
                //dr1["hovaten"] = hovaten;
                //dr1["chucvu"] = "";
                //dr1["phanloai"] = "";
                //dr1["ngaycongchuan"] = congtrongthang;
                //dr1["luongcb"] = luongcb;
                ////dr1["luongdongbh"] = 0;
                //dr1["tc_xangxe"] = pc_xangxe;
                //dr1["tc_trachnhiem"] = pc_trachnhiem;
                ////dr1["tongtrocap"] = 0;
                //dr1["tien_luong_kpi"] = luongkpi;
                //dr1["congthucte"] = congthucte;
                ////dr1["bosung"] = 0;
                //dr1["tongthunhap"] = 0;
                //dr1["tncn_banthan"] = TNCN_banthan;
                //dr1["tncn_songuoiphuthuoc"] = TNCN_nguoiphuthuoc;
                ////dr1["tncn_nguoiphuthuoc"] = 0;
                ////dr1["tncn_bhxh"] = 0;
                ////dr1["thunhaptinhthue"] = 0;
                ////dr1["thue_tncn"] = 0;
                ////dr1["dpcd"] = 0;
                ////dr1["thuclanh"] = 0;
                //dr1["tamungdot1"] = tamungdot1;
                ////dr1["conlai"] = 0;
                //dt.Rows.Add(dr1);

                var nRow = sheet.Rows[start_r];
                nRow.Cells[0].NumberValue = ++stt;
                nRow.Cells[1].Value = manv;
                nRow.Cells[2].Value = hovaten;
                nRow.Cells[3].Value = tenchucvu;
                nRow.Cells[4].Value = phanloai;
                nRow.Cells[5].NumberValue = (double)congtrongthang;
                nRow.Cells[6].NumberValue = (double)luongcb;
                nRow.Cells[16].NumberValue = (double)luongkpi;
                nRow.Cells[17].NumberValue = (double)congthucte;
                nRow.Cells[23].NumberValue = (double)TNCN_banthan;
                nRow.Cells[24].NumberValue = (double)TNCN_nguoiphuthuoc;

                sheet.CalculateAllValue();
                if (tamungdot1 > 0 && tamungdot1 < nRow.Cells[30].FormulaNumberValue)
                {
                    nRow.Cells[31].NumberValue = (double)tamungdot1;
                }

                ////Cập nhật database

                var person = SalaryUserModel.Where(d => d.person_id == item["NV_id"]).FirstOrDefault();

                person.MANV = manv;
                person.HOVATEN = hovaten;
                person.CHUCVU = tenchucvu;
                person.BOPHAN = phanloai;
                person.ngaycongchuan = congtrongthang;
                person.ngaycongthucte = congthucte;
                person.luongcb = (decimal)luongcb;
                person.luongdongbhxh = (decimal)Convert.ToSingle(nRow.Cells[7].FormulaValue);

                person.luongkpi = (decimal)luongkpi;
                person.tongthunhap = (decimal)Convert.ToSingle(nRow.Cells[19].FormulaValue);

                person.thunhapchiuthue = (decimal)Convert.ToSingle(nRow.Cells[22].FormulaValue);


                person.tncn_banthan = (decimal)TNCN_banthan;
                person.tncn_songuoiphuthuoc = (int)TNCN_nguoiphuthuoc;
                person.tncn_nguoiphuthuoc = (decimal)Convert.ToSingle(nRow.Cells[25].FormulaValue);
                person.tncn_bhxh = (decimal)Convert.ToSingle(nRow.Cells[26].FormulaValue);
                person.thunhaptinhthue = (decimal)Convert.ToSingle(nRow.Cells[27].FormulaValue);
                person.thue_tncn = (decimal)Convert.ToSingle(nRow.Cells[28].FormulaValue);
                person.dpcd = (decimal)Convert.ToSingle(nRow.Cells[29].FormulaValue);
                person.thuclanh = (decimal)Convert.ToSingle(nRow.Cells[30].FormulaValue);
                person.tamungdot1 = (decimal)Convert.ToSingle(nRow.Cells[31].FormulaValue);
                person.conlai = (decimal)Convert.ToSingle(nRow.Cells[32].FormulaValue);

                _context.Update(person);
                start_r++;

                //CellRange originDataRang = sheet.Range[$"A{(list_data_cong.Count() + 12)}:AZ{(list_data_cong.Count() + 12)}"];
                //CellRange targetDataRang = sheet.Range["A" + start_r + ":AZ" + start_r];
                //sheet.Copy(originDataRang, targetDataRang, true);

            }
            //sheet.InsertDataTable(dt, false, 13, 1);
            sheet.DeleteRow(12, 1);
            sheet.DeleteRow(start_r, 1);
            sheet.CalculateAllValue();
            workbook.SaveToFile(_configuration["Source:Path_Private"] + documentPath.Replace("/private", "").Replace("/", "\\"), ExcelVersion.Version2013);

            Model.file_url = documentPath;
            _context.Update(Model);
            _context.SaveChanges();
            //var congthuc_ct = _QLSXcontext.Congthuc_CTModel.Where()
            var jsonData = new { success = true };
            return Json(jsonData);

        }
        [HttpPost]
        public async Task<JsonResult> sendphieuluong(string id)
        {
            var SalaryUserModel = _context.SalaryUserModel.Where(d => d.salary_id == id).ToList();
            var Model = _context.SalaryModel.Where(d => d.id == id).FirstOrDefault();
            foreach (var s in SalaryUserModel)
            {
                var documentPath = _tinhcong.phieuluong(s.id, _configuration);
                var person = _context.PersonnelModel.Where(d => d.MANV == s.MANV).FirstOrDefault();

                string Domain = (HttpContext.Request.IsHttps ? "https://" : "http://") + HttpContext.Request.Host.Value;
                var body = _view.Render("Emails/Phieuluong", new
                {
                    link_logo = Domain + "/images/clientlogo_astahealthcare.com_f1800.png",
                    link = _configuration["Application:Info:link"] + "salary",
                    date_to = Model.date_to.Value
                });
                var attach = new List<string>()
                {
                    documentPath
                };
                var email = new EmailModel
                {
                    email_to = person.EMAIL,
                    subject = "[Phiếu lương] " + Model.name,
                    body = body,
                    email_type = "Phieuluong",
                    status = 1,
                    data_attachments = attach
                };
                _itContext.Add(email);
            }
            _itContext.SaveChanges();

            var jsonData = new { success = true };
            return Json(jsonData);

        }
        [HttpPost]
        public async Task<JsonResult> phieuluong(int id)
        {
            var documentPath = _tinhcong.phieuluong(id, _configuration);

            //var congthuc_ct = _QLSXcontext.Congthuc_CTModel.Where()
            var jsonData = new { success = true, link = documentPath };
            return Json(jsonData);

        }
        [HttpPost]
        public async Task<JsonResult> Save(SalaryModel SalaryModel, List<string> list_person)
        {
            System.Security.Claims.ClaimsPrincipal currentUser = this.User;
            var user_id = UserManager.GetUserId(currentUser);
            var user = await UserManager.GetUserAsync(currentUser);
            SalaryModel? SalaryModel_old;
            var properties = typeof(SalaryModel).GetProperties().Where(prop => prop.CanRead && prop.CanWrite && prop.PropertyType == typeof(DateTime?));

            foreach (var prop in properties)
            {
                DateTime? value = (DateTime?)prop.GetValue(SalaryModel, null);
                if (value != null && value.Value.Kind == DateTimeKind.Utc)
                {
                    value = value.Value.ToLocalTime();
                    prop.SetValue(SalaryModel, value, null);
                }
            }
            if (SalaryModel.id == null)
            {
                SalaryModel.id = Guid.NewGuid().ToString();
                SalaryModel.created_at = DateTime.Now;

                SalaryModel.created_by = user_id;


                _context.SalaryModel.Add(SalaryModel);

                _context.SaveChanges();

                SalaryModel_old = SalaryModel;

                //////
                var list_person_old = _context.SalaryUserModel.Where(d => d.salary_id == SalaryModel_old.id).Select(d => d.person_id).ToList();
                IEnumerable<string> list_delete_user = list_person_old.Except(list_person);
                IEnumerable<string> list_add_user = list_person.Except(list_person_old);

                if (list_add_user != null)
                {
                    foreach (string key in list_add_user)
                    {
                        var PersonModel = _context.PersonnelModel.Where(d => d.id == key).FirstOrDefault();

                        _context.Add(new SalaryUserModel()
                        {
                            salary_id = SalaryModel_old.id,
                            person_id = key,
                            email = PersonModel.EMAIL,

                        });
                    }
                    //_context.SaveChanges();
                }
                if (list_delete_user != null)
                {
                    foreach (string key in list_delete_user)
                    {
                        SalaryUserModel SalaryUserModel = _context.SalaryUserModel.Where(d => d.salary_id == SalaryModel_old.id && d.person_id == key).First();
                        _context.Remove(SalaryUserModel);
                    }
                    //_context.SaveChanges();
                }
                _context.SaveChanges();
            }
            else
            {

                SalaryModel_old = _context.SalaryModel.Where(d => d.id == SalaryModel.id).FirstOrDefault();
                CopyValues<SalaryModel>(SalaryModel_old, SalaryModel);
                SalaryModel_old.updated_at = DateTime.Now;

                _context.Update(SalaryModel_old);
                _context.SaveChanges();
            }

            return Json(new { success = true, data = SalaryModel_old });
        }

        [HttpPost]
        public async Task<JsonResult> Table()
        {
            var draw = Request.Form["draw"].FirstOrDefault();
            var start = Request.Form["start"].FirstOrDefault();
            var length = Request.Form["length"].FirstOrDefault();
            int pageSize = length != null ? Convert.ToInt32(length) : 0;
            var description = Request.Form["filters[description]"].FirstOrDefault();
            var name = Request.Form["filters[name]"].FirstOrDefault();
            var id = Request.Form["filters[id]"].FirstOrDefault();
            //var tenhh = Request.Form["filters[tenhh]"].FirstOrDefault();
            int skip = start != null ? Convert.ToInt32(start) : 0;
            var customerData = _context.SalaryModel.Where(d => d.deleted_at == null);
            int recordsTotal = customerData.Count();
            if (description != null && description != "")
            {
                customerData = customerData.Where(d => d.description.Contains(description));
            }
            if (name != null && name != "")
            {
                customerData = customerData.Where(d => d.name.Contains(name));
            }
            if (id != null)
            {
                customerData = customerData.Where(d => d.id == id);
            }

            int recordsFiltered = customerData.Count();
            var datapost = customerData.OrderByDescending(d => d.created_at).Skip(skip).Take(pageSize).ToList();
            //var data = new ArrayList();
            //foreach (var record in datapost)
            //{
            //	var ngaythietke = record.ngaythietke != null ? record.ngaythietke.Value.ToString("yyyy-MM-dd") : null;
            //	var ngaysodk = record.ngaysodk != null ? record.ngaysodk.Value.ToString("yyyy-MM-dd") : null;
            //	var ngayhethanthietke = record.ngayhethanthietke != null ? record.ngayhethanthietke.Value.ToString("yyyy-MM-dd") : null;
            //	var data1 = new
            //	{
            //		mahh = record.mahh,
            //		tenhh = record.tenhh,
            //		dvt = record.dvt,
            //		mansx = record.mansx,
            //		mancc = record.mancc,
            //		tennvlgoc = record.tennvlgoc,
            //		masothietke = record.masothietke,
            //		ghichu_thietke = record.ghichu_thietke,
            //		masodk = record.masodk,
            //		ghichu_sodk = record.ghichu_sodk,
            //		nhuongquyen = record.nhuongquyen,
            //		ngaythietke = ngaythietke,
            //		ngaysodk = ngaysodk,
            //		ngayhethanthietke = ngayhethanthietke
            //	};
            //	data.Add(data1);
            //}

            var jsonData = new { draw = draw, recordsFiltered = recordsFiltered, recordsTotal = recordsTotal, data = datapost };
            return Json(jsonData, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }

        public JsonResult Get(string id)
        {
            var data = _context.SalaryModel.Where(d => d.id == id).FirstOrDefault();
            return Json(data, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }
        public async Task<JsonResult> GetUser(string id)
        {
            var data = _context.SalaryUserModel.Where(d => d.salary_id == id).ToList();
            return Json(data, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }
        public async Task<JsonResult> listMySalary()
        {

            System.Security.Claims.ClaimsPrincipal currentUser = this.User;
            var user_id = UserManager.GetUserId(currentUser);
            var user = await UserManager.GetUserAsync(currentUser);
            var person = _context.PersonnelModel.Where(d => d.EMAIL == user.Email).FirstOrDefault();
            var MANV = person.MANV;

            var data = _context.SalaryUserModel.Where(d => d.MANV == MANV).Include(d => d.salary).Select(d => d.salary).ToList();
            data = data.Where(d => d.deleted_at == null).ToList();
            return Json(data, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }
        public async Task<JsonResult> GetMySalary(string id)
        {
            System.Security.Claims.ClaimsPrincipal currentUser = this.User;
            var user_id = UserManager.GetUserId(currentUser);
            var user = await UserManager.GetUserAsync(currentUser);
            var person = _context.PersonnelModel.Where(d => d.EMAIL == user.Email).FirstOrDefault();
            var MANV = person.MANV;


            var data = _context.SalaryUserModel.Where(d => d.salary_id == id && d.MANV == MANV).FirstOrDefault();
            return Json(data, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }

    }

}
