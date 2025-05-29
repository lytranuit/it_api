


using Holdtime.Models;
using Info.Models;
using it_template.Areas.Trend.Controllers;
using iText.Commons.Actions.Contexts;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Spire.Xls;
using System.Data;
using System.Text.Json.Serialization;
using Vue.Data;
using Vue.Models;
using Vue.Services;

namespace it_template.Areas.Info.Controllers
{

    //[Authorize(Roles = "Administrator,HR")]
    public class NguoiphuthuocController : BaseController
    {
        private readonly IConfiguration _configuration;
        private UserManager<UserModel> UserManager;
        public NguoiphuthuocController(NhansuContext context, AesOperation aes, IConfiguration configuration, UserManager<UserModel> UserMgr) : base(context, aes)
        {
            _configuration = configuration;
            UserManager = UserMgr;
        }

        [HttpPost]
        public async Task<JsonResult> Delete(int id)
        {
            var Model = _context.NguoiphuthuocModel.Where(d => d.id == id).FirstOrDefault();
            _context.Remove(Model);
            _context.SaveChanges();
            return Json(new { success = true, data = Model });
        }
        [HttpPost]
        public async Task<JsonResult> Save(List<NguoiphuthuocModel> list_add, List<NguoiphuthuocModel>? list_update, List<NguoiphuthuocModel>? list_delete)
        {
            if (list_delete != null)
                _context.RemoveRange(list_delete);
            if (list_add != null)
            {
                foreach (var item in list_add)
                {
                    _context.Add(item);
                }
            }
            if (list_update != null)
            {
                foreach (var item in list_update)
                {
                    _context.Update(item);
                }
            }

            _context.SaveChanges();


            return Json(new { success = true });
        }

        [HttpPost]
        public async Task<JsonResult> Table(string MANV)
        {
            var draw = Request.Form["draw"].FirstOrDefault();
            var start = Request.Form["start"].FirstOrDefault();
            var length = Request.Form["length"].FirstOrDefault();
            int pageSize = length != null ? Convert.ToInt32(length) : 0;
            var MAPHONG = Request.Form["filters[MAPHONG]"].FirstOrDefault();
            var TENPHONG = Request.Form["filters[TENPHONG]"].FirstOrDefault();
            var id = Request.Form["filters[id]"].FirstOrDefault();
            //var tenhh = Request.Form["filters[tenhh]"].FirstOrDefault();
            int skip = start != null ? Convert.ToInt32(start) : 0;
            var customerData = _context.NguoiphuthuocModel.Where(d => d.MANV == MANV);
            int recordsTotal = customerData.Count();
            int recordsFiltered = customerData.Count();
            var datapost = customerData.OrderBy(d => d.id).ToList();
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


        [HttpPost]
        public async Task<JsonResult> excel()
        {
            var viewPath = "wwwroot/report/excel/Danhsachnguoiphuthuoc.xlsx";
            var documentPath = "/private/info/data/" + DateTime.Now.ToFileTimeUtc() + ".xlsx";

            string Domain = (HttpContext.Request.IsHttps ? "https://" : "http://") + HttpContext.Request.Host.Value;

            Workbook workbook = new Workbook();
            workbook.LoadFromFile(viewPath);

            var datapost = _context.NguoiphuthuocModel.Include(d => d.PersonnelModel).Where(d => d.PersonnelModel != null && d.PersonnelModel.NGAYNGHIVIEC == null).ToList();
            Worksheet sheet = workbook.Worksheets[0];
            int start_r = 1; // Dòng gốc để sao chép định dạng

            sheet.InsertRow(start_r + 1, datapost.Count(), InsertOptionsType.FormatAsAfter);
            foreach (var record in datapost)
            {
                var nRow = sheet.Rows[start_r];
                nRow.Cells[0].Value = record.MANV.Trim();
                nRow.Cells[1].Value = record.PersonnelModel.HOVATEN.Trim();
                nRow.Cells[2].Value = record.tennguoiphuthuoc.Trim();
                nRow.Cells[3].Value = record.namsinh;
                nRow.Cells[4].Value = record.quanhe;
                nRow.Cells[5].Value = record.masothue ?? "";
                nRow.Cells[6].Value = record.is_phuthuoc == true ? "Phụ thuộc" : "Người thân";

                start_r++;
            }

            workbook.SaveToFile(_configuration["Source:Path_Private"] + documentPath.Replace("/private", "").Replace("/", "\\"), ExcelVersion.Version2013);

            //var congthuc_ct = _QLSXcontext.Congthuc_CTModel.Where()
            var jsonData = new { success = true, link = documentPath };
            return Json(jsonData);
        }


        public async Task<JsonResult> import()
        {
            return Json(new { });
            var nguoithan = _context.NguoiphuthuocModel.FromSqlRaw("SELECT CAST(ROW_NUMBER() OVER (ORDER BY MANV) AS INT) AS ID,  MANV, QUANHE,HOTEN as tennguoiphuthuoc, CONVERT(VARCHAR, NGAYSINH, 103) AS namsinh,CAST(0 AS BIT) as is_phuthuoc,'' as masothue  FROM nsQUANHEGIADINH").AsNoTracking().ToList();
            foreach (var item in nguoithan)
            {
                item.id = 0;
                item.masothue = null;
            }
            _context.AddRange(nguoithan);
            _context.SaveChanges();

            return Json(new { });
        }
    }

}
