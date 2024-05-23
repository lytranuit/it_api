

using elFinder.NetCore.Models;
using iText.Commons.Actions.Contexts;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Spire.Xls;
using Spire.Xls.Core;
using System.Collections;
using System.Data;
using System.Drawing;
using System.Runtime.Intrinsics.X86;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using Vue.Data;
using Vue.Models;

namespace it_template.Areas.Trend.Controllers
{

    public class LimitController : BaseController
    {
        private readonly IConfiguration _configuration;
        private UserManager<UserModel> UserManager;
        public LimitController(ItContext context, IConfiguration configuration, UserManager<UserModel> UserMgr) : base(context)
        {
            _configuration = configuration;
            UserManager = UserMgr;
        }

        [HttpPost]
        public async Task<JsonResult> Delete(int id)
        {
            var Model = _context.LimitModel.Where(d => d.id == id).FirstOrDefault();
            Model.deleted_at = DateTime.Now;
            _context.Update(Model);
            _context.SaveChanges();
            return Json(new { success = true, data = Model });
        }
        [HttpPost]
        public async Task<JsonResult> Save(LimitModel LimitModel, List<int> list_point)
        {
            System.Security.Claims.ClaimsPrincipal currentUser = this.User;
            var user_id = UserManager.GetUserId(currentUser);
            var user = await UserManager.GetUserAsync(currentUser);
            if (LimitModel.date_effect != null && LimitModel.date_effect.Value.Kind == DateTimeKind.Utc)
            {
                LimitModel.date_effect = LimitModel.date_effect.Value.ToLocalTime();
            }
            if (LimitModel.date_from != null && LimitModel.date_from.Value.Kind == DateTimeKind.Utc)
            {
                LimitModel.date_from = LimitModel.date_from.Value.ToLocalTime();
            }
            if (LimitModel.date_to != null && LimitModel.date_to.Value.Kind == DateTimeKind.Utc)
            {
                LimitModel.date_to = LimitModel.date_to.Value.ToLocalTime();
            }
            LimitModel.points = null;
            LimitModel? LimitModel_old;
            if (LimitModel.id == 0)
            {
                LimitModel.created_at = DateTime.Now;


                _context.LimitModel.Add(LimitModel);

                _context.SaveChanges();

                LimitModel_old = LimitModel;

            }
            else
            {

                LimitModel_old = _context.LimitModel.Where(d => d.id == LimitModel.id).FirstOrDefault();
                CopyValues<LimitModel>(LimitModel_old, LimitModel);
                LimitModel_old.updated_at = DateTime.Now;

                _context.Update(LimitModel_old);
                _context.SaveChanges();
            }


            var LimitPointModel_old = _context.LimitPointModel.Where(d => d.limit_id == LimitModel_old.id).ToList();
            _context.RemoveRange(LimitPointModel_old);
            _context.SaveChanges();

            foreach (var item in list_point)
            {
                _context.Add(new LimitPointModel()
                {
                    limit_id = LimitModel_old.id,
                    point_id = item
                });
            }
            _context.SaveChanges();

            return Json(new { success = true, data = LimitModel_old });
        }
        [HttpPost]
        public async Task<JsonResult> Table()
        {
            var draw = Request.Form["draw"].FirstOrDefault();
            var start = Request.Form["start"].FirstOrDefault();
            var length = Request.Form["length"].FirstOrDefault();
            int pageSize = length != null ? Convert.ToInt32(length) : 0;
            var name = Request.Form["filters[name]"].FirstOrDefault();
            var object_id_text = Request.Form["filters[object_id]"].FirstOrDefault();
            var target_id_text = Request.Form["filters[target_id]"].FirstOrDefault();
            var id_text = Request.Form["filters[id]"].FirstOrDefault();
            int id = id_text != null ? Convert.ToInt32(id_text) : 0;
            int object_id = object_id_text != null ? Convert.ToInt32(object_id_text) : 0;
            int target_id = target_id_text != null ? Convert.ToInt32(target_id_text) : 0;
            //var tenhh = Request.Form["filters[tenhh]"].FirstOrDefault();
            int skip = start != null ? Convert.ToInt32(start) : 0;
            var customerData = _context.LimitModel.Where(d => d.deleted_at == null);
            int recordsTotal = customerData.Count();
            if (name != null && name != "")
            {
                customerData = customerData.Where(d => d.name.Contains(name));
            }
            if (id != 0)
            {
                customerData = customerData.Where(d => d.id == id);
            }
            if (object_id != 0)
            {
                customerData = customerData.Where(d => d.object_id == object_id);
            }
            if (target_id != 0)
            {
                customerData = customerData.Where(d => d.target_id == target_id);
            }
            int recordsFiltered = customerData.Count();
            var datapost = customerData.OrderByDescending(d => d.date_effect).ThenByDescending(d => d.id).Skip(skip).Take(pageSize).Include(d => d.target).Include(d => d.obj).ToList();
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
            return Json(jsonData);
        }
        [HttpPost]
        public async Task<JsonResult> xuatexcel(LimitModel LimitModel, List<int> list_point)
        {
            DateTime? date_from = null;
            DateTime? date_to = null;
            if (LimitModel.date_from != null && LimitModel.date_from.Value.Kind == DateTimeKind.Utc)
            {
                date_from = LimitModel.date_from.Value.ToLocalTime();
            }
            if (LimitModel.date_to != null && LimitModel.date_to.Value.Kind == DateTimeKind.Utc)
            {
                date_to = LimitModel.date_to.Value.ToLocalTime();
            }
            if (LimitModel.date_effect != null && LimitModel.date_effect.Value.Kind == DateTimeKind.Utc)
            {
                LimitModel.date_effect = LimitModel.date_effect.Value.ToLocalTime();
            }
            ////Lấy 
            var list = _context.ResultModel
                .Where(d => list_point.Contains(d.point_id.Value) && d.deleted_at == null && d.date >= date_from && d.date <= date_to)
                .Include(d => d.point).ThenInclude(d => d.location).OrderBy(d => d.date).ToList();

            var data = new ArrayList();

            var viewPath = _configuration["Source:Path_Private"] + "\\trend\\templates\\GHCB_GHHD.xls";
            var documentPath = "/temp/" + DateTime.Now.ToFileTimeUtc() + ".xlsx";
            string Domain = (HttpContext.Request.IsHttps ? "https://" : "http://") + HttpContext.Request.Host.Value;

            Workbook workbook = new Workbook();
            workbook.LoadFromFile(viewPath);

            //Create a font
            ExcelFont font1 = workbook.CreateFont();
            font1.FontName = "Times New Roman";
            font1.IsBold = false;
            font1.Size = 10;

            //Create another font
            ExcelFont font2 = workbook.CreateFont();
            font2.IsBold = false;
            font2.IsItalic = true;
            font2.FontName = "Times New Roman";
            font2.Size = 10;

            Worksheet sheet1 = workbook.Worksheets[0];
            var targetModel = _context.TargetModel.Where(d => d.id == LimitModel.target_id).FirstOrDefault();

            var cell = sheet1.Range["C5"];
            cell.Value = LimitModel.standard_limit.Value.ToString();


            RichText richText1 = sheet1.Range["C4"].RichText;
            richText1.Text = targetModel.name + "\r\n" + targetModel.name_en;
            richText1.SetFont(0, targetModel.name.Length, font1);

            richText1.SetFont(targetModel.name.Length, richText1.Text.Length, font2);


            var cell2 = sheet1.Range["G4"];
            cell2.Value = date_from.Value.ToString("dd/MM/yyyy") + " - " + date_to.Value.ToString("dd/MM/yyyy");

            var cell3 = sheet1.Range["D25"];
            cell3.Value = LimitModel.date_effect.Value.ToString("dd/MM/yyyy");
            //list = list.Where(d => d.deleted_at == null).Distinct().ToList();
            if (list.Count > 0)
            {
                Worksheet sheet = workbook.Worksheets[0];
                int stt = 0;
                var start_r = 9;

                DataTable dt = new DataTable();
                dt.Columns.Add("stt", typeof(int));
                dt.Columns.Add("ngay", typeof(string));
                dt.Columns.Add("department", typeof(string));
                dt.Columns.Add("department_temp1", typeof(string));
                dt.Columns.Add("department_temp2", typeof(string));
                dt.Columns.Add("point_code", typeof(string));
                dt.Columns.Add("value", typeof(decimal));
                sheet.InsertRow(9, list.Count(), InsertOptionsType.FormatAsBefore);
                foreach (var item in list)
                {
                    DataRow dr1 = dt.NewRow();
                    dr1["stt"] = (++stt);
                    dr1["ngay"] = item.date.Value.ToString("dd/MM/yyyy");
                    //var richtext = new RichText(item.location.name);
                    dr1["department"] = item.point.location.name + "\r\n" + item.point.location.name_en;
                    dr1["point_code"] = item.point.code;
                    dr1["value"] = item.value;
                    dt.Rows.Add(dr1);
                    start_r++;

                }
                sheet.InsertDataTable(dt, false, 9, 1);
                sheet.DeleteRow(8);
                start_r = 8;


                foreach (var item in list)
                {
                    //Console.WriteLine(start_r);
                    //Console.WriteLine("Lenght:" + item.point.location.name.Length);
                    RichText richText = sheet.Range["C" + start_r].RichText;
                    richText.SetFont(0, item.point.location.name.Length, font1);

                    richText.SetFont(item.point.location.name.Length, richText.Text.Length, font2);
                    start_r++;

                    //CellRange originDataRang = sheet.Range["A8:G8"];
                    //CellRange targetDataRang = sheet.Range["A" + start_r + ":G" + start_r];
                    //sheet.Copy(originDataRang, targetDataRang, true);
                }
                //avg = sheet.Range["C" + (start_r + 3)].FormulaNumberValue;
                //sheet.CalculateAllValue();
                //var avg = sheet.Range["C" + (start_r + 3)].FormulaNumberValue;
                //var min = sheet.Range["C" + (start_r + 4)].FormulaNumberValue;
                //var max = sheet.Range["C" + (start_r + 5)].FormulaNumberValue;

            }

            workbook.SaveToFile("./wwwroot" + documentPath, ExcelVersion.Version2013);
            return Json(new { success = true, link = Domain + documentPath });
        }
        public JsonResult Get(int id)
        {
            var data = _context.LimitModel.Where(d => d.id == id).Include(d => d.points).FirstOrDefault();
            return Json(data);
        }
        private void CopyValues<T>(T target, T source)
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
