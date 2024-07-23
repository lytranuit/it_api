using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.CodeAnalysis;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Doc.Reporting;
using Spire.Xls;
using System.Collections;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Text.Json.Serialization;
using Vue.Data;
using Vue.Models;
using Vue.Services;

namespace it_template.Areas.Trend.Controllers
{

    public class ReportController : BaseController
    {
        private readonly IConfiguration _configuration;
        private UserManager<UserModel> UserManager;
        private AesOperation _AesOperation;
        public ReportController(ItContext context, IConfiguration configuration, UserManager<UserModel> UserMgr, AesOperation aes) : base(context)
        {
            _configuration = configuration;
            UserManager = UserMgr;
            _AesOperation = aes;
        }

        [HttpPost]
        public async Task<JsonResult> xuatlimit(LimitModel LimitModel, List<int> list_point)
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
                .Where(d => list_point.Contains(d.point_id.Value) && d.deleted_at == null && d.date >= date_from && d.date <= date_to && d.target_id == LimitModel.target_id && d.object_id == LimitModel.object_id)
                .Include(d => d.point).ThenInclude(d => d.location).OrderBy(d => d.date).ToList();

            var data = new ArrayList();

            var targetModel = _context.TargetModel.Where(d => d.id == LimitModel.target_id).FirstOrDefault();




            if (LimitModel.object_id == 2)
            {
                var viewPath = _configuration["Source:Path_Private"] + "\\trend\\templates\\GHCB_GHHD(Visinh).xls";
                var documentPath = "/temp/GHCB_GHHD(Visinh)_" + DateTime.Now.ToFileTimeUtc() + ".xlsx";
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

                var cell = sheet1.Range["C5"];
                cell.Value = LimitModel.standard_limit.Value.ToString();


                RichText richText1 = sheet1.Range["C4"].RichText;
                richText1.Text = targetModel.name + "\r\n" + targetModel.name_en;
                richText1.SetFont(0, targetModel.name.Length, font1);

                richText1.SetFont(targetModel.name.Length, richText1.Text.Length, font2);


                var cell2 = sheet1.Range["G4"];
                cell2.Value = date_from.Value.ToString("dd/MM/yyyy") + " - " + date_to.Value.ToString("dd/MM/yyyy");

                //var cell3 = sheet1.Range["D25"];
                //cell3.Value = LimitModel.date_effect.Value.ToString("dd/MM/yyyy");
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
            else
            {
                var viewPath = _configuration["Source:Path_Private"] + "\\trend\\templates\\GHCB_GHHD(Nuoc).xls";
                var documentPath = "/temp/GHCB_GHHD(Nuoc)_" + DateTime.Now.ToFileTimeUtc() + ".xlsx";
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
                var cell = sheet1.Range["C5"];
                cell.Value = LimitModel.standard_limit.Value.ToString();


                RichText richText1 = sheet1.Range["C4"].RichText;
                richText1.Text = targetModel.name + "\r\n" + targetModel.name_en;
                richText1.SetFont(0, targetModel.name.Length, font1);

                richText1.SetFont(targetModel.name.Length, richText1.Text.Length, font2);

                var cell3 = sheet1.Range["G4"];
                cell3.Value = targetModel.unit;


                var cell2 = sheet1.Range["G5"];
                cell2.Value = date_from.Value.ToString("dd/MM/yyyy") + " - " + date_to.Value.ToString("dd/MM/yyyy");

                //var cell3 = sheet1.Range["D25"];
                //cell3.Value = LimitModel.date_effect.Value.ToString("dd/MM/yyyy");
                //list = list.Where(d => d.deleted_at == null).Distinct().ToList();
                if (list.Count > 0)
                {
                    Worksheet sheet = workbook.Worksheets[0];
                    int stt = 0;
                    var start_r = 9;

                    DataTable dt = new DataTable();
                    dt.Columns.Add("stt", typeof(int));
                    dt.Columns.Add("ngay", typeof(string));
                    dt.Columns.Add("point_code", typeof(string));
                    dt.Columns.Add("point_code_temp1", typeof(string));
                    dt.Columns.Add("point_code_temp2", typeof(string));
                    dt.Columns.Add("point_code_temp3", typeof(string));
                    dt.Columns.Add("value", typeof(decimal));
                    sheet.InsertRow(9, list.Count(), InsertOptionsType.FormatAsBefore);
                    foreach (var item in list)
                    {
                        DataRow dr1 = dt.NewRow();
                        dr1["stt"] = (++stt);
                        dr1["ngay"] = item.date.Value.ToString("dd/MM/yyyy");
                        dr1["point_code"] = item.point.code;
                        dr1["value"] = item.value;
                        dt.Rows.Add(dr1);
                        start_r++;

                    }
                    sheet.InsertDataTable(dt, false, 9, 1);
                    sheet.DeleteRow(8);
                    //avg = sheet.Range["C" + (start_r + 3)].FormulaNumberValue;
                    //sheet.CalculateAllValue();
                    //var avg = sheet.Range["C" + (start_r + 3)].FormulaNumberValue;
                    //var min = sheet.Range["C" + (start_r + 4)].FormulaNumberValue;
                    //var max = sheet.Range["C" + (start_r + 5)].FormulaNumberValue;

                }
                workbook.SaveToFile("./wwwroot" + documentPath, ExcelVersion.Version2013);

                return Json(new { success = true, link = Domain + documentPath, data = list });
            }

            return Json(new { success = false });

        }

        //[HttpPost]
        //public async Task<JsonResult> xuatraw1(LimitModel LimitModel, List<int> list_point)
        //{
        //    DateTime? date_from = null;
        //    DateTime? date_to = null;
        //    if (LimitModel.date_from != null && LimitModel.date_from.Value.Kind == DateTimeKind.Utc)
        //    {
        //        date_from = LimitModel.date_from.Value.ToLocalTime();
        //    }
        //    if (LimitModel.date_to != null && LimitModel.date_to.Value.Kind == DateTimeKind.Utc)
        //    {
        //        date_to = LimitModel.date_to.Value.ToLocalTime();
        //    }
        //    if (LimitModel.date_effect != null && LimitModel.date_effect.Value.Kind == DateTimeKind.Utc)
        //    {
        //        LimitModel.date_effect = LimitModel.date_effect.Value.ToLocalTime();
        //    }
        //    ////Lấy 
        //    var list = _context.ResultModel
        //        .Where(d => list_point.Contains(d.point_id.Value) && d.deleted_at == null && d.date >= date_from && d.date <= date_to && d.target_id == LimitModel.target_id && d.object_id == LimitModel.object_id)
        //        .Include(d => d.point).ThenInclude(d => d.location).OrderBy(d => d.date).ToList();

        //    var data = new ArrayList();

        //    var targetModel = _context.TargetModel.Where(d => d.id == LimitModel.target_id).FirstOrDefault();

        //    var labels = list.Select(d => d.point.code).Distinct().OrderBy(d => d).ToList();
        //    var list_data = list.GroupBy(d => d.date).Select(d => new
        //    {
        //        date = d.Key,
        //        data = d.ToList(),
        //    }).OrderBy(d => d.date).ToList();
        //    if (LimitModel.object_id == 2)
        //    {
        //        var viewPath = _configuration["Source:Path_Private"] + "\\trend\\templates\\dulieu(Visinh).xlsx";
        //        var documentPath = "/temp/dulieu(Visinh)_" + DateTime.Now.ToFileTimeUtc() + ".xlsx";
        //        string Domain = (HttpContext.Request.IsHttps ? "https://" : "http://") + HttpContext.Request.Host.Value;

        //        Workbook workbook = new Workbook();
        //        workbook.LoadFromFile(viewPath);

        //        //Create a font
        //        ExcelFont font1 = workbook.CreateFont();
        //        font1.FontName = "Times New Roman";
        //        font1.IsBold = false;
        //        font1.Size = 18;

        //        //Create another font
        //        ExcelFont font2 = workbook.CreateFont();
        //        font2.IsBold = false;
        //        font2.IsItalic = true;
        //        font2.FontName = "Times New Roman";
        //        font2.Size = 18;

        //        Worksheet sheet1 = workbook.Worksheets[0];


        //        RichText richText1 = sheet1.Range["A1"].RichText;
        //        richText1.Text = targetModel.name + "\r\n" + targetModel.name_en;
        //        richText1.SetFont(0, targetModel.name.Length, font1);

        //        richText1.SetFont(targetModel.name.Length - 1, richText1.Text.Length, font2);


        //        if (list_data.Count > 0)
        //        {
        //            Worksheet sheet = workbook.Worksheets[0];
        //            int stt = 0;
        //            var start_r = 5;

        //            DataTable dt = new DataTable();
        //            dt.Columns.Add("stt", typeof(int));
        //            dt.Columns.Add("ngay", typeof(string));
        //            sheet.InsertColumn(4, labels.Count(), InsertOptionsType.FormatAsBefore);
        //            var rowlocation = sheet1.Rows[1];
        //            var rowpoint = sheet1.Rows[2];
        //            var stt_cell = 2;

        //            var locations = list.GroupBy(d => d.point.location.name).Select(d => new
        //            {
        //                location = d.Key,
        //                points = d.GroupBy(e => e.point.code).Select(e => new
        //                {
        //                    point = e.Key,
        //                    data = e.ToList()
        //                }).OrderBy(e => e.point).ToList(),
        //            }).OrderBy(d => d.location).ToList();

        //            foreach (var location in locations)
        //            {
        //                var cell = rowlocation.Cells[stt_cell];
        //                var columnname = ColumnIndexToColumnLetter(stt_cell + 1);
        //                var columnname2 = ColumnIndexToColumnLetter(stt_cell + location.points.Count());

        //                cell.Value = location.location;

        //                var range = sheet.Range[columnname + "2:" + columnname2 + "2"];
        //                range.Merge();
        //                foreach (var point in location.points)
        //                {
        //                    var cell1 = rowpoint.Cells[stt_cell++];
        //                    cell1.Value = point.point;

        //                    dt.Columns.Add("value_" + point.point, typeof(decimal));
        //                }
        //            }




        //            sheet.InsertRow(5, list_data.Count(), InsertOptionsType.FormatAsBefore);
        //            foreach (var item in list_data)
        //            {
        //                DataRow dr1 = dt.NewRow();
        //                dr1["stt"] = (++stt);
        //                dr1["ngay"] = item.date.Value.ToString("dd/MM/yyyy");

        //                foreach (var location in locations)
        //                {
        //                    foreach (var point in location.points)
        //                    {
        //                        var d = item.data.Where(d => d.point.code == point.point).FirstOrDefault();
        //                        if (d != null)
        //                        {
        //                            dr1["value_" + point.point] = d.value;
        //                        }
        //                    }
        //                }
        //                dt.Rows.Add(dr1);
        //                start_r++;

        //            }
        //            sheet.InsertDataTable(dt, false, 5, 1);
        //            sheet.DeleteRow(4);
        //            //sheet.DeleteColumn(3);
        //            //avg = sheet.Range["C" + (start_r + 3)].FormulaNumberValue;
        //            //sheet.CalculateAllValue();
        //            //var avg = sheet.Range["C" + (start_r + 3)].FormulaNumberValue;
        //            //var min = sheet.Range["C" + (start_r + 4)].FormulaNumberValue;
        //            //var max = sheet.Range["C" + (start_r + 5)].FormulaNumberValue;

        //        }
        //        workbook.SaveToFile("./wwwroot" + documentPath, ExcelVersion.Version2013);

        //        return Json(new { success = true, link = Domain + documentPath, data = list });
        //    }
        //    else
        //    {
        //        var viewPath = _configuration["Source:Path_Private"] + "\\trend\\templates\\dulieu(Nuoc).xlsx";
        //        var documentPath = "/temp/dulieu(Nuoc)_" + DateTime.Now.ToFileTimeUtc() + ".xlsx";
        //        string Domain = (HttpContext.Request.IsHttps ? "https://" : "http://") + HttpContext.Request.Host.Value;

        //        Workbook workbook = new Workbook();
        //        workbook.LoadFromFile(viewPath);

        //        //Create a font
        //        ExcelFont font1 = workbook.CreateFont();
        //        font1.FontName = "Times New Roman";
        //        font1.IsBold = false;
        //        font1.Size = 18;

        //        //Create another font
        //        ExcelFont font2 = workbook.CreateFont();
        //        font2.IsBold = false;
        //        font2.IsItalic = true;
        //        font2.FontName = "Times New Roman";
        //        font2.Size = 18;

        //        Worksheet sheet1 = workbook.Worksheets[0];


        //        RichText richText1 = sheet1.Range["A1"].RichText;
        //        richText1.Text = targetModel.name + "\r\n" + targetModel.name_en;
        //        richText1.SetFont(0, targetModel.name.Length, font1);

        //        richText1.SetFont(targetModel.name.Length, richText1.Text.Length, font2);


        //        if (list_data.Count > 0)
        //        {
        //            Worksheet sheet = workbook.Worksheets[0];
        //            int stt = 0;
        //            var start_r = 4;

        //            DataTable dt = new DataTable();
        //            dt.Columns.Add("stt", typeof(int));
        //            dt.Columns.Add("ngay", typeof(string));
        //            sheet.InsertColumn(4, labels.Count(), InsertOptionsType.FormatAsBefore);
        //            var row = sheet1.Rows[1];
        //            var stt_cell = 2;
        //            foreach (var label in labels)
        //            {
        //                var cell = row.Cells[stt_cell++];
        //                cell.Value = label;

        //                dt.Columns.Add("value_" + label, typeof(decimal));
        //            }
        //            sheet.InsertRow(4, list_data.Count(), InsertOptionsType.FormatAsBefore);
        //            foreach (var item in list_data)
        //            {
        //                DataRow dr1 = dt.NewRow();
        //                dr1["stt"] = (++stt);
        //                dr1["ngay"] = item.date.Value.ToString("dd/MM/yyyy");

        //                foreach (var label in labels)
        //                {
        //                    var d = item.data.Where(d => d.point.code == label).FirstOrDefault();
        //                    if (d != null)
        //                    {
        //                        dr1["value_" + label] = d.value;
        //                    }
        //                }
        //                dt.Rows.Add(dr1);
        //                start_r++;

        //            }
        //            sheet.InsertDataTable(dt, false, 4, 1);
        //            sheet.DeleteRow(3);
        //            //sheet.DeleteColumn(3);
        //            //avg = sheet.Range["C" + (start_r + 3)].FormulaNumberValue;
        //            //sheet.CalculateAllValue();
        //            //var avg = sheet.Range["C" + (start_r + 3)].FormulaNumberValue;
        //            //var min = sheet.Range["C" + (start_r + 4)].FormulaNumberValue;
        //            //var max = sheet.Range["C" + (start_r + 5)].FormulaNumberValue;

        //        }
        //        workbook.SaveToFile("./wwwroot" + documentPath, ExcelVersion.Version2013);

        //        return Json(new { success = true, link = Domain + documentPath, data = list });
        //    }

        //    return Json(new { success = false });

        //}

        [HttpPost]
        public async Task<JsonResult> xuatraw(LimitModel LimitModel, List<int> list_point)
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
                .Where(d => list_point.Contains(d.point_id.Value) && d.deleted_at == null && d.date >= date_from && d.date <= date_to && d.object_id == LimitModel.object_id)
                .Include(d => d.point).ThenInclude(d => d.location)
                .Include(d => d.target)
                .OrderBy(d => d.date).ToList();

            var data = new ArrayList();

            var documentPath = "/temp/Rawdata_" + DateTime.Now.ToFileTimeUtc() + ".xlsx";
            string Domain = (HttpContext.Request.IsHttps ? "https://" : "http://") + HttpContext.Request.Host.Value;
            if (LimitModel.object_id == 2)
            {
                var viewPath = _configuration["Source:Path_Private"] + "\\trend\\templates\\0000052_A03_02 - Raw data for of microbial results of (room, equiment) quality monitoring_effec. 24.05.24_issue1.xlsx";
                Workbook workbook = new Workbook();
                workbook.LoadFromFile(viewPath);

                //Create a font
                ExcelFont font1 = workbook.CreateFont();
                font1.FontName = "Times New Roman";
                font1.IsBold = false;
                font1.Size = 18;

                //Create another font
                ExcelFont font2 = workbook.CreateFont();
                font2.IsBold = false;
                font2.IsItalic = true;
                font2.FontName = "Times New Roman";
                font2.Size = 18;

                Worksheet sheet1 = workbook.Worksheets[0];




                if (list.Count > 0)
                {
                    Worksheet sheet = workbook.Worksheets[0];
                    int stt = 0;
                    var start_r = 5;

                    DataTable dt = new DataTable();
                    //dt.Columns.Add("stt", typeof(int));
                    dt.Columns.Add("ngay", typeof(string));
                    dt.Columns.Add("department", typeof(string));
                    dt.Columns.Add("area", typeof(string));
                    dt.Columns.Add("tansuat", typeof(string));
                    dt.Columns.Add("point_code", typeof(string));
                    dt.Columns.Add("target", typeof(string));
                    dt.Columns.Add("value", typeof(string));

                    var rowlocation = sheet1.Rows[1];
                    var rowpoint = sheet1.Rows[2];
                    var stt_cell = 2;



                    sheet.InsertRow(5, list.Count(), InsertOptionsType.FormatAsBefore);
                    foreach (var item in list)
                    {
                        var capsach = _context.LocationModel.Where(d => d.id == item.point.location.parent).FirstOrDefault();
                        var khuvuc = _context.LocationModel.Where(d => d.id == capsach.parent).FirstOrDefault();
                        var area = khuvuc != null ? capsach.name + "-" + khuvuc.name : capsach.name;
                        var value_d = (double)item.value;
                        var value = value_d.ToString("#,##0.##");
                        if (item.target.value_type == "boolean")
                        {
                            if (item.value == 1)
                            {
                                value = item.target.text_yes;
                            }
                            else
                            {
                                value = item.target.text_no;
                            }
                        }
                        else if (item.target.id == 11)
                        {
                            value = value_d.ToString("#,##0.0#");
                        }
                        DataRow dr1 = dt.NewRow();
                        //dr1["stt"] = (++stt);
                        dr1["ngay"] = item.date.Value.ToString("dd/MM/yyyy");
                        dr1["department"] = item.point.location.name;
                        dr1["area"] = area;
                        dr1["tansuat"] = item.point.frequency;
                        dr1["point_code"] = item.point.code;
                        dr1["target"] = item.target.name;
                        dr1["value"] = value;

                        dt.Rows.Add(dr1);
                        start_r++;

                    }
                    sheet.InsertDataTable(dt, false, 5, 1);
                    sheet.DeleteRow(4);
                    //sheet.DeleteColumn(3);
                    //avg = sheet.Range["C" + (start_r + 3)].FormulaNumberValue;
                    //sheet.CalculateAllValue();
                    //var avg = sheet.Range["C" + (start_r + 3)].FormulaNumberValue;
                    //var min = sheet.Range["C" + (start_r + 4)].FormulaNumberValue;
                    //var max = sheet.Range["C" + (start_r + 5)].FormulaNumberValue;

                }
                workbook.SaveToFile("./wwwroot" + documentPath, ExcelVersion.Version2013);

            }
            else if (LimitModel.object_id == 3 || LimitModel.object_id == 5 || LimitModel.object_id == 6)
            {
                var viewPath = _configuration["Source:Path_Private"] + "\\trend\\templates\\0000052_A02_02 - Raw data for Water system quality monitoring_effec. 24.05.24_issue1.xlsx";
                Workbook workbook = new Workbook();
                workbook.LoadFromFile(viewPath);

                //Create a font
                ExcelFont font1 = workbook.CreateFont();
                font1.FontName = "Times New Roman";
                font1.IsBold = false;
                font1.Size = 18;

                //Create another font
                ExcelFont font2 = workbook.CreateFont();
                font2.IsBold = false;
                font2.IsItalic = true;
                font2.FontName = "Times New Roman";
                font2.Size = 18;

                Worksheet sheet1 = workbook.Worksheets[0];




                if (list.Count > 0)
                {
                    Worksheet sheet = workbook.Worksheets[0];
                    int stt = 0;
                    var start_r = 5;

                    DataTable dt = new DataTable();
                    //dt.Columns.Add("stt", typeof(int));
                    dt.Columns.Add("ngay", typeof(string));
                    dt.Columns.Add("department", typeof(string));
                    dt.Columns.Add("area", typeof(string));
                    dt.Columns.Add("tansuat", typeof(string));
                    dt.Columns.Add("point_code", typeof(string));
                    dt.Columns.Add("target", typeof(string));
                    dt.Columns.Add("value", typeof(string));

                    var rowlocation = sheet1.Rows[1];
                    var rowpoint = sheet1.Rows[2];
                    var stt_cell = 2;



                    sheet.InsertRow(5, list.Count(), InsertOptionsType.FormatAsBefore);
                    foreach (var item in list)
                    {
                        var capsach = _context.LocationModel.Where(d => d.id == item.point.location.parent).FirstOrDefault();
                        var khuvuc = _context.LocationModel.Where(d => d.id == capsach.parent).FirstOrDefault();
                        var area = khuvuc != null ? capsach.name + "-" + khuvuc.name : capsach.name;
                        var value_d = (double)item.value;
                        var value = value_d.ToString("#,##0.##");
                        if (item.target.value_type == "boolean")
                        {
                            if (item.value == 1)
                            {
                                value = item.target.text_yes;
                            }
                            else
                            {
                                value = item.target.text_no;
                            }
                        }
                        else if (item.target.id == 11)
                        {
                            value = value_d.ToString("#,##0.0#");
                        }
                        DataRow dr1 = dt.NewRow();
                        //dr1["stt"] = (++stt);
                        dr1["ngay"] = item.date.Value.ToString("dd/MM/yyyy");
                        dr1["area"] = item.point.location.name;
                        dr1["department"] = area;
                        dr1["tansuat"] = item.point.frequency;
                        dr1["point_code"] = item.point.code;
                        dr1["target"] = item.target.name;
                        dr1["value"] = value;

                        dt.Rows.Add(dr1);
                        start_r++;

                    }
                    sheet.InsertDataTable(dt, false, 5, 1);
                    sheet.DeleteRow(4);
                    //sheet.DeleteColumn(3);
                    //avg = sheet.Range["C" + (start_r + 3)].FormulaNumberValue;
                    //sheet.CalculateAllValue();
                    //var avg = sheet.Range["C" + (start_r + 3)].FormulaNumberValue;
                    //var min = sheet.Range["C" + (start_r + 4)].FormulaNumberValue;
                    //var max = sheet.Range["C" + (start_r + 5)].FormulaNumberValue;

                }
                workbook.SaveToFile("./wwwroot" + documentPath, ExcelVersion.Version2013);

            }
            else if (LimitModel.object_id == 4)
            {
                var viewPath = _configuration["Source:Path_Private"] + "\\trend\\templates\\0000052_A01_02 - Raw data for CA system quality monitoring_effec. 24.05.24_issue1.xlsx";
                Workbook workbook = new Workbook();
                workbook.LoadFromFile(viewPath);

                //Create a font
                ExcelFont font1 = workbook.CreateFont();
                font1.FontName = "Times New Roman";
                font1.IsBold = false;
                font1.Size = 18;

                //Create another font
                ExcelFont font2 = workbook.CreateFont();
                font2.IsBold = false;
                font2.IsItalic = true;
                font2.FontName = "Times New Roman";
                font2.Size = 18;

                Worksheet sheet1 = workbook.Worksheets[0];




                if (list.Count > 0)
                {
                    Worksheet sheet = workbook.Worksheets[0];
                    int stt = 0;
                    var start_r = 5;

                    DataTable dt = new DataTable();
                    //dt.Columns.Add("stt", typeof(int));
                    dt.Columns.Add("ngay", typeof(string));
                    dt.Columns.Add("department", typeof(string));
                    dt.Columns.Add("area", typeof(string));
                    dt.Columns.Add("tansuat", typeof(string));
                    dt.Columns.Add("point_code", typeof(string));
                    dt.Columns.Add("target", typeof(string));
                    dt.Columns.Add("value", typeof(string));

                    var rowlocation = sheet1.Rows[1];
                    var rowpoint = sheet1.Rows[2];
                    var stt_cell = 2;



                    sheet.InsertRow(5, list.Count(), InsertOptionsType.FormatAsBefore);
                    foreach (var item in list)
                    {
                        var capsach = _context.LocationModel.Where(d => d.id == item.point.location.parent).FirstOrDefault();
                        var khuvuc = _context.LocationModel.Where(d => d.id == capsach.parent).FirstOrDefault();
                        var area = khuvuc != null ? capsach.name + "-" + khuvuc.name : capsach.name;
                        var value_d = (double)item.value;
                        var value = value_d.ToString("#,##0.##");
                        if (item.target.value_type == "boolean")
                        {
                            if (item.value == 1)
                            {
                                value = item.target.text_yes;
                            }
                            else
                            {
                                value = item.target.text_no;
                            }
                        }
                        else if (item.target.id == 11)
                        {
                            value = value_d.ToString("#,##0.0#");
                        }
                        DataRow dr1 = dt.NewRow();
                        //dr1["stt"] = (++stt);
                        dr1["ngay"] = item.date.Value.ToString("dd/MM/yyyy");
                        dr1["area"] = item.point.location.name;
                        dr1["department"] = area;
                        dr1["tansuat"] = item.point.frequency;
                        dr1["point_code"] = item.point.code;
                        dr1["target"] = item.target.name;
                        dr1["value"] = value;

                        dt.Rows.Add(dr1);
                        start_r++;

                    }
                    sheet.InsertDataTable(dt, false, 5, 1);
                    sheet.DeleteRow(4);
                    //sheet.DeleteColumn(3);
                    //avg = sheet.Range["C" + (start_r + 3)].FormulaNumberValue;
                    //sheet.CalculateAllValue();
                    //var avg = sheet.Range["C" + (start_r + 3)].FormulaNumberValue;
                    //var min = sheet.Range["C" + (start_r + 4)].FormulaNumberValue;
                    //var max = sheet.Range["C" + (start_r + 5)].FormulaNumberValue;

                }
                workbook.SaveToFile("./wwwroot" + documentPath, ExcelVersion.Version2013);

            }


            return Json(new { success = true, link = Domain + documentPath, data = list });

        }

        [HttpPost]
        public async Task<JsonResult> xuattrend(LimitModel LimitModel, List<int> list_point, int timestamp, int type_bc, DateTime? date_from_prev, DateTime? date_to_prev)
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
            if (date_from_prev != null && date_from_prev.Value.Kind == DateTimeKind.Utc)
            {
                date_from_prev = date_from_prev.Value.ToLocalTime();
            }
            if (date_to_prev != null && date_to_prev.Value.Kind == DateTimeKind.Utc)
            {
                date_to_prev = date_to_prev.Value.ToLocalTime();
            }
            var data = new ArrayList();


            if (LimitModel.object_id == 2)
            {
                var viewPath = _configuration["Source:Path_Private"] + "\\trend\\templates\\template_trend(Visinh).docx";
                var documentPath = "/temp/trend(Visinh)_" + DateTime.Now.ToFileTimeUtc() + ".docx";
                string Domain = (HttpContext.Request.IsHttps ? "https://" : "http://") + HttpContext.Request.Host.Value;

                var raw = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    {"dates",LimitModel.date_from.Value.ToString("dd/MM/yyyy") + " - " + LimitModel.date_to.Value.ToString("dd/MM/yyyy") },
                };
                ///XUẤT PDF

                //Creates Document instance
                Spire.Doc.Document document = new Spire.Doc.Document();
                Section section = document.AddSection();

                document.LoadFromFile(viewPath, Spire.Doc.FileFormat.Docx);

                string[] fieldName = raw.Keys.ToArray();
                string[] fieldValue = raw.Values.ToArray();

                string[] MergeFieldNames = document.MailMerge.GetMergeFieldNames();
                string[] GroupNames = document.MailMerge.GetMergeGroupNames();

                document.MailMerge.Execute(fieldName, fieldValue);


                //document.MailMerge.ExecuteWidthRegion(datatable_details);

                ////Lấy 
                ///

                var from = date_from_prev != null ? date_from_prev : date_from;
                var results = _context.ResultModel.Where(d => d.deleted_at == null && list_point.Contains(d.point_id.Value) && d.date >= from && d.date <= date_to)
                   .Include(d => d.target)
                   .Include(d => d.point)
                   .ThenInclude(d => d.location).ToList();

                var list = results
                    .Where(d => list_point.Contains(d.point_id.Value) && d.deleted_at == null && d.date >= date_from && d.date <= date_to && d.object_id == LimitModel.object_id)
                    .OrderBy(d => d.date).ToList();

                var list_prev = results
                  .Where(d => list_point.Contains(d.point_id.Value) && d.deleted_at == null && d.date >= date_from_prev && d.date <= date_to_prev && d.object_id == LimitModel.object_id)
                  .OrderBy(d => d.date).ToList();

                var tables = new List<Label1>();
                var groups = list
                         .GroupBy(f => f.target).Select(f => new Label1()
                         {
                             key = f.Key.id,
                             text = f.Key.name,
                             timestamp = new DateTimeOffset(DateTime.UtcNow).ToUnixTimeSeconds(),
                             data = f.ToList(),
                         }).ToList();
                Dictionary<string, Spire.Doc.Table> replacements_table = new Dictionary<string, Spire.Doc.Table>(StringComparer.OrdinalIgnoreCase) { };


                /////Table
                foreach (var group in groups)
                {
                    var labels = group.data.GroupBy(d => d.point).Select(d => new
                    {
                        date = d.OrderBy(d => d.date).First().date,
                        id = d.Key.id,
                        point_code = d.Key.code,
                        department = d.Key.location.name,
                        department_en = d.Key.location.name_en
                    }).OrderBy(d => d.date).ThenBy(d => d.point_code).ToList();
                    var batchSize = 10;
                    var processed = 0;
                    var hasNextBatch = true;
                    var stt = 0;
                    while (hasNextBatch)
                    {
                        var batch = labels.Skip(processed).Take(batchSize).ToList();
                        if (batch.Count == 0)
                        {
                            hasNextBatch = false;
                            continue;
                        }
                        var list_batch = batch.Select(d => d.id).ToList();
                        var date_batch = group.data.Where(d => list_batch.Contains(d.point_id.Value)).GroupBy(d => d.date).Select(d => d.Key).ToList();
                        var text = "table_" + group.key + "_" + stt++;

                        //if (replacements_table.Count() == 25)
                        //{

                        //document = new Spire.Doc.Document();
                        //section = document.AddSection();

                        //document.LoadFromFile(viewPath, Spire.Doc.FileFormat.Docx);
                        //}

                        Spire.Doc.Table table1 = section.AddTable(true);
                        PreferredWidth width = new PreferredWidth(WidthType.Percentage, 100);
                        table1.PreferredWidth = width;
                        replacements_table.Add(text, table1);
                        table1.ResetCells(8 + date_batch.Count(), batch.Count + 2);

                        var row_current = 0;
                        TableRow FRow;
                        var i = 0;
                        Paragraph p1;
                        TextRange TR1;
                        if (row_current == 0)
                        {
                            //// Row 0
                            //Set the first row as table header
                            FRow = table1.Rows[row_current];

                            FRow.IsHeader = true;
                            //Set the height and color of the first row
                            FRow.Height = 30;
                            i = 0;

                            table1.ApplyHorizontalMerge(row_current, 0, 1);

                            FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                            p1 = FRow.Cells[i].AddParagraph();
                            p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                            TR1 = p1.AppendText("Tên phòng/ thiết bị");

                            //TR1.CharacterFormat.FontName = "Arial";

                            TR1.CharacterFormat.FontSize = 10;

                            TR1.CharacterFormat.Bold = true;

                            p1 = FRow.Cells[i].AddParagraph();
                            p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                            TR1 = p1.AppendText("Room/ equipment name");

                            //TR1.CharacterFormat.FontName = "Arial";

                            TR1.CharacterFormat.FontSize = 10;

                            TR1.CharacterFormat.Italic = true;
                            i++;

                            ////
                            i++;

                            var department_cu = "";
                            foreach (var column in batch)
                            {
                                if (department_cu != column.department)
                                {

                                    var count_department = batch.Where(d => d.department == column.department).Count();
                                    if (count_department > 1)
                                    {
                                        table1.ApplyHorizontalMerge(row_current, i, i + count_department - 1);
                                    }
                                    department_cu = column.department;

                                    //Set alignment for cells

                                    FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                                    Paragraph p = FRow.Cells[i].AddParagraph();

                                    p.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                    //Set data format
                                    TextRange TR = p.AppendText(column.department);

                                    //TR.CharacterFormat.FontName = "Arial";

                                    TR.CharacterFormat.FontSize = 10;

                                    //TR.CharacterFormat.Bold = true;
                                    p = FRow.Cells[i].AddParagraph();

                                    p.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                    //Set data format
                                    TR = p.AppendText(column.department_en);
                                    TR.CharacterFormat.FontSize = 10;
                                    TR.CharacterFormat.Italic = true;
                                }
                                i++;
                            }

                            row_current++;
                        }

                        if (row_current == 1)
                        {
                            //// Row 1
                            //Set the first row as table header
                            FRow = table1.Rows[row_current];

                            FRow.IsHeader = true;
                            //Set the height and color of the first row
                            FRow.Height = 30;
                            i = 0;

                            table1.ApplyHorizontalMerge(1, 0, 1);

                            p1 = FRow.Cells[i].AddParagraph();
                            FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                            p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                            TR1 = p1.AppendText("Vị trí lấy mẫu");

                            //TR1.CharacterFormat.FontName = "Arial";

                            TR1.CharacterFormat.FontSize = 10;

                            TR1.CharacterFormat.Bold = true;

                            p1 = FRow.Cells[i].AddParagraph();
                            p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                            TR1 = p1.AppendText("Sampling location");

                            //TR1.CharacterFormat.FontName = "Arial";

                            TR1.CharacterFormat.FontSize = 10;

                            TR1.CharacterFormat.Italic = true;
                            i++;

                            ////
                            i++;


                            foreach (var column in batch)
                            {
                                //Set alignment for cells

                                Paragraph p = FRow.Cells[i].AddParagraph();

                                FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                                p.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                //Set data format
                                TextRange TR = p.AppendText(column.point_code);

                                //TR.CharacterFormat.FontName = "Arial";

                                TR.CharacterFormat.FontSize = 10;

                                //TR.CharacterFormat.Bold = true;
                                i++;
                            }
                            row_current++;

                        }

                        if (row_current == 2)
                        {
                            //// Row 2
                            //Set the first row as table header
                            FRow = table1.Rows[row_current];

                            //FRow.IsHeader = true;
                            //Set the height and color of the first row
                            FRow.Height = 30;
                            i = 0;

                            p1 = FRow.Cells[i].AddParagraph();
                            FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                            p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                            TR1 = p1.AppendText("Stt");

                            //TR1.CharacterFormat.FontName = "Arial";

                            TR1.CharacterFormat.FontSize = 10;


                            p1 = FRow.Cells[i].AddParagraph();
                            p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                            TR1 = p1.AppendText("No");

                            //TR1.CharacterFormat.FontName = "Arial";

                            TR1.CharacterFormat.FontSize = 10;

                            TR1.CharacterFormat.Italic = true;
                            //TR1.CharacterFormat.Bold = true;
                            i++;
                            p1 = FRow.Cells[i].AddParagraph();
                            FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                            p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                            TR1 = p1.AppendText("Ngày");
                            TR1.CharacterFormat.FontSize = 10;


                            p1 = FRow.Cells[i].AddParagraph();
                            p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                            TR1 = p1.AppendText("Date");

                            //TR1.CharacterFormat.FontName = "Arial";

                            TR1.CharacterFormat.FontSize = 10;

                            TR1.CharacterFormat.Italic = true;
                            ////
                            i++;


                            table1.ApplyHorizontalMerge(2, 2, batch.Count + 1);
                            p1 = FRow.Cells[i].AddParagraph();
                            FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                            p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                            TR1 = p1.AppendText("Kết quả / ");
                            TR1.CharacterFormat.FontSize = 10;
                            TR1 = p1.AppendText("Result: ");
                            TR1.CharacterFormat.FontSize = 10;
                            TR1.CharacterFormat.Italic = true;
                            TR1 = p1.AppendText("CFU / Plate");
                            TR1.CharacterFormat.FontSize = 10;
                            TR1.CharacterFormat.Bold = true;
                            row_current++;

                        }
                        if ("data" == "data")
                        {

                            var stt_table = 0;
                            foreach (var date in date_batch)
                            {
                                stt_table++;
                                //// Row 0
                                //Set the first row as table header
                                FRow = table1.Rows[row_current];

                                //FRow.IsHeader = true;
                                //Set the height and color of the first row
                                FRow.Height = 30;
                                i = 0;


                                p1 = FRow.Cells[i].AddParagraph();
                                FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                                p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                TR1 = p1.AppendText(stt_table.ToString());

                                //TR1.CharacterFormat.FontName = "Arial";

                                TR1.CharacterFormat.FontSize = 10;

                                //TR1.CharacterFormat.Bold = true;
                                i++;

                                ////

                                p1 = FRow.Cells[i].AddParagraph();
                                FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                                p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                TR1 = p1.AppendText(date.Value.ToString("dd/MM/yyyy"));

                                //TR1.CharacterFormat.FontName = "Arial";

                                TR1.CharacterFormat.FontSize = 10;

                                //TR1.CharacterFormat.Bold = true;
                                i++;


                                foreach (var column in batch)
                                {
                                    //Set alignment for cells
                                    var data_point = group.data.Where(d => d.point_id == column.id && d.date == date).FirstOrDefault();
                                    if (data_point != null)
                                    {
                                        Paragraph p = FRow.Cells[i].AddParagraph();

                                        FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                                        p.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                        //Set data format
                                        TextRange TR = p.AppendText(data_point.value.Value.ToString("#,##0.##"));

                                        //TR.CharacterFormat.FontName = "Arial";

                                        TR.CharacterFormat.FontSize = 10;

                                        //TR.CharacterFormat.Bold = true;

                                    }
                                    else
                                    {
                                        FRow.Cells[i].CellFormat.BackColor = Color.LightGray;
                                    }
                                    i++;
                                }

                                row_current++;
                            }
                        }

                        if ("Max" == "Max")
                        {
                            //// Row Max
                            //Set the first row as table header
                            FRow = table1.Rows[row_current];

                            //FRow.IsHeader = true;
                            //Set the height and color of the first row
                            FRow.Height = 30;
                            i = 0;

                            table1.ApplyHorizontalMerge(row_current, 0, 1);

                            p1 = FRow.Cells[i].AddParagraph();
                            FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                            p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                            TR1 = p1.AppendText("Max");

                            //TR1.CharacterFormat.FontName = "Arial";

                            TR1.CharacterFormat.FontSize = 10;

                            TR1.CharacterFormat.Bold = true;
                            i++;

                            ////
                            i++;


                            foreach (var column in batch)
                            {
                                //Set alignment for cells
                                var data_point = group.data.Where(d => d.point_id == column.id).Select(d => d.value).Max();
                                if (data_point != null)
                                {
                                    Paragraph p = FRow.Cells[i].AddParagraph();

                                    FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                                    p.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                    //Set data format
                                    TextRange TR = p.AppendText(data_point.Value.ToString("#,##0.##"));

                                    //TR.CharacterFormat.FontName = "Arial";

                                    TR.CharacterFormat.FontSize = 10;

                                    //TR.CharacterFormat.Bold = true;

                                }
                                else
                                {
                                    FRow.Cells[i].CellFormat.BackColor = Color.LightGray;
                                }
                                i++;
                            }
                            row_current++;

                        }
                        if ("Min" == "Min")
                        {
                            //// Row Min
                            //Set the first row as table header
                            FRow = table1.Rows[row_current];

                            //FRow.IsHeader = true;
                            //Set the height and color of the first row
                            FRow.Height = 30;
                            i = 0;

                            table1.ApplyHorizontalMerge(row_current, 0, 1);

                            p1 = FRow.Cells[i].AddParagraph();
                            FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                            p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                            TR1 = p1.AppendText("Min");

                            //TR1.CharacterFormat.FontName = "Arial";

                            TR1.CharacterFormat.FontSize = 10;

                            TR1.CharacterFormat.Bold = true;
                            i++;

                            ////
                            i++;


                            foreach (var column in batch)
                            {
                                //Set alignment for cells
                                var data_point = group.data.Where(d => d.point_id == column.id).Select(d => d.value).Min();
                                if (data_point != null)
                                {
                                    Paragraph p = FRow.Cells[i].AddParagraph();

                                    FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                                    p.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                    //Set data format
                                    TextRange TR = p.AppendText(data_point.Value.ToString("#,##0.##"));

                                    //TR.CharacterFormat.FontName = "Arial";

                                    TR.CharacterFormat.FontSize = 10;

                                    //TR.CharacterFormat.Bold = true;

                                }
                                else
                                {
                                    FRow.Cells[i].CellFormat.BackColor = Color.LightGray;
                                }
                                i++;
                            }
                            row_current++;

                        }

                        if (1 == 1)
                        {
                            //// Row Kết quả trước đó
                            //Set the first row as table header
                            FRow = table1.Rows[row_current];

                            //FRow.IsHeader = true;
                            //Set the height and color of the first row
                            FRow.Height = 30;
                            i = 0;

                            table1.ApplyHorizontalMerge(row_current, 0, batch.Count() + 1);

                            p1 = FRow.Cells[i].AddParagraph();
                            FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                            p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                            TR1 = p1.AppendText("Kết quả trước đó / ");

                            //TR1.CharacterFormat.FontName = "Arial";

                            TR1.CharacterFormat.FontSize = 10;

                            TR1.CharacterFormat.Bold = true;

                            //p1 = FRow.Cells[i].AddParagraph();
                            //p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                            TR1 = p1.AppendText("Results of previous");

                            //TR1.CharacterFormat.FontName = "Arial";

                            TR1.CharacterFormat.FontSize = 10;

                            TR1.CharacterFormat.Italic = true;
                            i++;
                            row_current++;
                        }

                        if ("Max_prev" == "Max_prev")
                        {
                            //// Row Max Trước đó
                            //Set the first row as table header
                            FRow = table1.Rows[row_current];

                            //FRow.IsHeader = true;
                            //Set the height and color of the first row
                            FRow.Height = 30;
                            i = 0;

                            table1.ApplyHorizontalMerge(row_current, 0, 1);

                            p1 = FRow.Cells[i].AddParagraph();
                            FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                            p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                            TR1 = p1.AppendText("Max");

                            //TR1.CharacterFormat.FontName = "Arial";

                            TR1.CharacterFormat.FontSize = 10;

                            TR1.CharacterFormat.Bold = true;
                            i++;

                            ////
                            i++;


                            foreach (var column in batch)
                            {
                                //Set alignment for cells
                                var data_point = list_prev.Where(d => d.point_id == column.id).Select(d => d.value).Max();
                                if (data_point != null)
                                {
                                    Paragraph p = FRow.Cells[i].AddParagraph();

                                    FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                                    p.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                    //Set data format
                                    TextRange TR = p.AppendText(data_point.Value.ToString("#,##0.##"));

                                    //TR.CharacterFormat.FontName = "Arial";

                                    TR.CharacterFormat.FontSize = 10;

                                    //TR.CharacterFormat.Bold = true;

                                }
                                else
                                {
                                    FRow.Cells[i].CellFormat.BackColor = Color.LightGray;
                                }
                                i++;
                            }
                            row_current++;

                        }
                        if ("Min_prev" == "Min_prev")
                        {
                            //// Row Min Trước đó
                            //Set the first row as table header
                            FRow = table1.Rows[row_current];

                            //FRow.IsHeader = true;
                            //Set the height and color of the first row
                            FRow.Height = 30;
                            i = 0;

                            table1.ApplyHorizontalMerge(row_current, 0, 1);

                            p1 = FRow.Cells[i].AddParagraph();
                            FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                            p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                            TR1 = p1.AppendText("Min");

                            //TR1.CharacterFormat.FontName = "Arial";

                            TR1.CharacterFormat.FontSize = 10;

                            TR1.CharacterFormat.Bold = true;
                            i++;

                            ////
                            i++;



                            foreach (var column in batch)
                            {
                                //Set alignment for cells
                                var data_point = list_prev.Where(d => d.point_id == column.id).Select(d => d.value).Min();
                                if (data_point != null)
                                {
                                    Paragraph p = FRow.Cells[i].AddParagraph();

                                    FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                                    p.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                    //Set data format
                                    TextRange TR = p.AppendText(data_point.Value.ToString("#,##0.##"));

                                    //TR.CharacterFormat.FontName = "Arial";

                                    TR.CharacterFormat.FontSize = 10;

                                    //TR.CharacterFormat.Bold = true;

                                }
                                else
                                {
                                    FRow.Cells[i].CellFormat.BackColor = Color.LightGray;
                                }
                                i++;
                            }
                            row_current++;
                        }




                        tables.Add(new Label1()
                        {
                            key = group.key,
                            table_ketqua = text
                        });
                        processed += batch.Count;
                        hasNextBatch = batch.Count == batchSize;
                    }

                }

                //Trend
                var groups_trend = results.Where(d => d.target.value_type == "float")
                               .GroupBy(e => new { e.point.frequency_id, e.point.frequency, e.point.frequency_en }).Select(e => new Label1()
                               {
                                   key = e.Key.frequency_id.Value,
                                   label = e.Key.frequency,
                                   label_en = e.Key.frequency_en,
                                   children = e
                                   .GroupBy(f => f.target).Select(f => new Label1()
                                   {
                                       key = f.Key.id,
                                       label = f.Key.name,
                                       label_en = f.Key.name_en,
                                       parent = e.Key.frequency_id.Value,
                                       children = f.GroupBy(d => d.point.location).Select(d => new Label1()
                                       {
                                           key = d.Key.id,
                                           label = d.Key.name,
                                           label_en = d.Key.name_en,
                                           parent = f.Key.id,
                                       }).ToList(),
                                   }).ToList()
                               }).ToList();
                var fres = new List<Label1>();
                var targets = new List<Label1>();
                var locations = new List<Label1>();

                foreach (var f in groups_trend)
                {
                    fres.Add(f);
                    foreach (var t in f.children)
                    {
                        targets.Add(t);
                        foreach (var l in t.children)
                        {
                            l.image_ketqua = $"chart_{f.key}_{t.key}_{l.key}";
                            l.image_link_ketqua = $"link_{f.key}_{t.key}_{l.key}";
                            var chartData = _context.ChartDataModel.Where(d => d.timestamp == timestamp && d.key == l.image_ketqua).FirstOrDefault();
                            if (chartData != null)
                            {
                                l.link_ketqua = _configuration["Application:Trend:link"] + "report/chart/" + Uri.EscapeDataString(_AesOperation.EncryptString(chartData.id.ToString()));
                            }
                            locations.Add(l);
                        }
                    }
                }


                List<DictionaryEntry> relationsList = new List<DictionaryEntry>();
                relationsList.Add(new DictionaryEntry("groups", string.Empty));
                relationsList.Add(new DictionaryEntry("children", "key = %groups.key%"));

                //relationsList.Add(new DictionaryEntry("fres", string.Empty));
                //relationsList.Add(new DictionaryEntry("targets", "parent = %fres.key%"));
                //relationsList.Add(new DictionaryEntry("locations", "parent = %locations.key%"));

                MailMergeDataSet mailMergeDataSet = new MailMergeDataSet();
                mailMergeDataSet.Add(new MailMergeDataTable("groups", groups));
                mailMergeDataSet.Add(new MailMergeDataTable("children", tables));

                //mailMergeDataSet.Add(new MailMergeDataTable("fres", fres));
                //mailMergeDataSet.Add(new MailMergeDataTable("targets", targets));
                //mailMergeDataSet.Add(new MailMergeDataTable("locations", locations));

                document.MailMerge.ExecuteWidthNestedRegion(mailMergeDataSet, relationsList);


                foreach (KeyValuePair<string, Spire.Doc.Table> entry in replacements_table)
                {
                    TextSelection selection = document.FindString(entry.Key, true, true);
                    if (selection == null)
                    {
                        var table = entry.Value;
                        //table
                        continue;
                    }
                    TextRange range = selection.GetAsOneRange();
                    Paragraph paragraph = range.OwnerParagraph;
                    Body body = paragraph.OwnerTextBody;
                    int index = body.ChildObjects.IndexOf(paragraph);
                    //var fontsize = paragrap

                    body.ChildObjects.Remove(paragraph);
                    body.ChildObjects.Insert(index, entry.Value);
                }

                relationsList = new List<DictionaryEntry>();
                //relationsList.Add(new DictionaryEntry("groups", string.Empty));
                //relationsList.Add(new DictionaryEntry("children", "key = %groups.key%"));

                relationsList.Add(new DictionaryEntry("fres", string.Empty));
                relationsList.Add(new DictionaryEntry("targets", "parent = %fres.key%"));
                relationsList.Add(new DictionaryEntry("locations", "parent = %targets.key%"));

                mailMergeDataSet = new MailMergeDataSet();
                //mailMergeDataSet.Add(new MailMergeDataTable("groups", groups));
                //mailMergeDataSet.Add(new MailMergeDataTable("children", tables));

                mailMergeDataSet.Add(new MailMergeDataTable("fres", fres));
                mailMergeDataSet.Add(new MailMergeDataTable("targets", targets));
                mailMergeDataSet.Add(new MailMergeDataTable("locations", locations));

                document.MailMerge.ExecuteWidthNestedRegion(mailMergeDataSet, relationsList);
                //Set Font Style and Size

                ParagraphStyle style = new ParagraphStyle(document);

                style.Name = "FontStyle";
                style.CharacterFormat.FontSize = 12;

                document.Styles.Add(style);

                foreach (var location in locations)
                {
                    TextSelection selection = document.FindString(location.image_ketqua, true, true);
                    TextSelection selection1 = document.FindString(location.image_link_ketqua, true, true);
                    if (selection == null)
                    {
                        //var table = entry.Value;
                        //table
                        continue;
                    }
                    Image image = Image.FromFile($"./wwwroot/tmp/{timestamp}/{location.image_ketqua}.png");
                    Paragraph paragraph = section.AddParagraph();

                    DocPicture pic = new DocPicture(document);
                    pic.LoadImage(image);
                    pic.Width = 750;
                    pic.Height = 350;
                    paragraph.AppendHyperlink(location.link_ketqua, "Xem biểu đồ", HyperlinkType.WebLink);
                    paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center;

                    paragraph.ApplyStyle(style.Name);

                    var range = selection.GetAsOneRange();
                    var index = range.OwnerParagraph.ChildObjects.IndexOf(range);
                    range.OwnerParagraph.ChildObjects.Insert(index, pic);
                    range.OwnerParagraph.ChildObjects.Remove(range);

                    TextRange range1 = selection1.GetAsOneRange();
                    Paragraph paragraph1 = range1.OwnerParagraph;
                    Body body = paragraph1.OwnerTextBody;
                    int index1 = body.ChildObjects.IndexOf(paragraph1);
                    //var fontsize = paragrap

                    body.ChildObjects.Remove(paragraph1);
                    body.ChildObjects.Insert(index1, paragraph);
                    //var range1 = selection1.GetAsOneRange();
                    //var index1 = range1.OwnerParagraph.ChildObjects.IndexOf(range1);
                    //range1.OwnerParagraph.ChildObjects.Insert(index1, paragraph);
                    //range1.OwnerParagraph.ChildObjects.Remove(range1);

                }

                document.SaveToFile("./wwwroot" + documentPath, Spire.Doc.FileFormat.Docx);

                return Json(new { success = true, link = Domain + documentPath, });
            }
            else
            {
                var viewPath = _configuration["Source:Path_Private"] + "\\trend\\templates\\template_trend(Nuoc).docx";
                var documentPath = "/temp/trend_" + DateTime.Now.ToFileTimeUtc() + ".docx";
                string Domain = (HttpContext.Request.IsHttps ? "https://" : "http://") + HttpContext.Request.Host.Value;

                var raw = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    {"dates",LimitModel.date_from.Value.ToString("dd/MM/yyyy") + " - " + LimitModel.date_to.Value.ToString("dd/MM/yyyy") },
                };
                ///XUẤT PDF

                //Creates Document instance
                Spire.Doc.Document document = new Spire.Doc.Document();
                Section section = document.AddSection();

                document.LoadFromFile(viewPath, Spire.Doc.FileFormat.Docx);

                string[] fieldName = raw.Keys.ToArray();
                string[] fieldValue = raw.Values.ToArray();

                string[] MergeFieldNames = document.MailMerge.GetMergeFieldNames();
                string[] GroupNames = document.MailMerge.GetMergeGroupNames();

                document.MailMerge.Execute(fieldName, fieldValue);


                ////Lấy 
                ///

                var from = date_from_prev != null ? date_from_prev : date_from;
                var results = _context.ResultModel.Where(d => d.deleted_at == null && list_point.Contains(d.point_id.Value) && d.date >= from && d.date <= date_to)
                   .Include(d => d.target)
                   .Include(d => d.point).ToList();
                var list = results
                    .Where(d => list_point.Contains(d.point_id.Value) && d.deleted_at == null && d.date >= date_from && d.date <= date_to && d.object_id == LimitModel.object_id)
                    .OrderBy(d => d.date).ToList();

                var list_prev = results
                  .Where(d => list_point.Contains(d.point_id.Value) && d.deleted_at == null && d.date >= date_from_prev && d.date <= date_to_prev && d.object_id == LimitModel.object_id)
                  .OrderBy(d => d.date).ToList();
                Dictionary<string, Spire.Doc.Table> replacements_table = new Dictionary<string, Spire.Doc.Table>(StringComparer.OrdinalIgnoreCase) { };
                var tables = new List<Label1>();
                var points = list.GroupBy(d => d.point_id.Value).Select(d => new Points
                {
                    id = d.Key,
                    point = d.First().point.code,
                    data = d.ToList(),
                }).ToList();
                var stt = 0;
                foreach (var point in points)
                {
                    //var data_point_prev = list_prev.Where(d => d.point_id == point.id).ToList();
                    var batch = point.data.GroupBy(d => d.target).Select(d => new
                    {
                        id = d.Key.id,
                        target = d.Key,
                    }).OrderBy(d => d.id).ToList();
                    var date_batch = point.data.GroupBy(d => d.date).Select(d => d.Key).ToList();
                    var text = "table_" + point.id + "_" + stt++;


                    Spire.Doc.Table table1 = section.AddTable(true);
                    PreferredWidth width = new PreferredWidth(WidthType.Percentage, 100);
                    table1.PreferredWidth = width;
                    replacements_table.Add(text, table1);
                    table1.ResetCells(7 + date_batch.Count(), batch.Count + 1);

                    var row_current = 0;
                    TableRow FRow;
                    var i = 0;
                    Paragraph p1;
                    TextRange TR1;
                    if (row_current == 0)
                    {
                        //// Row 1
                        //Set the first row as table header
                        FRow = table1.Rows[row_current];

                        FRow.IsHeader = true;
                        //Set the height and color of the first row
                        FRow.Height = 30;
                        i = 0;


                        table1.ApplyHorizontalMerge(row_current, 0, batch.Count());

                        FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                        FRow.Cells[i].CellFormat.BackColor = Color.LightGray;
                        p1 = FRow.Cells[i].AddParagraph();
                        p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                        TR1 = p1.AppendText(point.point);

                        //TR1.CharacterFormat.FontName = "Arial";

                        TR1.CharacterFormat.FontSize = 10;

                        TR1.CharacterFormat.Bold = true;

                        i++;




                        row_current++;

                    }

                    if (row_current == 1)
                    {
                        //// Row 1
                        //Set the first row as table header
                        FRow = table1.Rows[row_current];

                        FRow.IsHeader = true;
                        //Set the height and color of the first row
                        FRow.Height = 30;
                        i = 0;


                        FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                        FRow.Cells[i].CellFormat.BackColor = Color.LightGray;
                        p1 = FRow.Cells[i].AddParagraph();
                        p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                        TR1 = p1.AppendText("Ngày");

                        //TR1.CharacterFormat.FontName = "Arial";

                        TR1.CharacterFormat.FontSize = 10;

                        TR1.CharacterFormat.Bold = true;

                        p1 = FRow.Cells[i].AddParagraph();
                        p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                        TR1 = p1.AppendText("Date");

                        //TR1.CharacterFormat.FontName = "Arial";

                        TR1.CharacterFormat.FontSize = 10;

                        TR1.CharacterFormat.Italic = true;
                        i++;



                        foreach (var column in batch)
                        {
                            //Set alignment for cells

                            Paragraph p = FRow.Cells[i].AddParagraph();

                            FRow.Cells[i].CellFormat.BackColor = Color.LightGray;
                            FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                            p.Format.HorizontalAlignment = HorizontalAlignment.Center;
                            //Set data format
                            TextRange TR = p.AppendText(column.target.name);

                            //TR.CharacterFormat.FontName = "Arial";

                            TR.CharacterFormat.FontSize = 10;

                            //TR.CharacterFormat.Bold = true;

                            p = FRow.Cells[i].AddParagraph();

                            p.Format.HorizontalAlignment = HorizontalAlignment.Center;
                            //Set data format
                            TR = p.AppendText(column.target.name_en);
                            TR.CharacterFormat.FontSize = 10;
                            TR.CharacterFormat.Italic = true;
                            i++;
                        }
                        row_current++;

                    }


                    if ("data" == "data")
                    {

                        var stt_table = 0;
                        foreach (var date in date_batch)
                        {
                            stt_table++;
                            //// Row 0
                            //Set the first row as table header
                            FRow = table1.Rows[row_current];

                            //FRow.IsHeader = true;
                            //Set the height and color of the first row
                            FRow.Height = 30;
                            i = 0;

                            ////

                            p1 = FRow.Cells[i].AddParagraph();
                            FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                            p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                            TR1 = p1.AppendText(date.Value.ToString("dd/MM/yyyy"));

                            //TR1.CharacterFormat.FontName = "Arial";

                            TR1.CharacterFormat.FontSize = 10;

                            //TR1.CharacterFormat.Bold = true;
                            i++;


                            foreach (var column in batch)
                            {
                                //Set alignment for cells
                                var data_point = point.data.Where(d => d.target_id == column.id && d.date == date).FirstOrDefault();
                                if (data_point != null)
                                {

                                    if (column.target.value_type == "boolean")
                                    {
                                        var value_vi = data_point.value == 1 ? column.target.text_yes : column.target.text_no;
                                        var value_en = data_point.value == 1 ? column.target.text_yes_en : column.target.text_no_en;
                                        Paragraph p = FRow.Cells[i].AddParagraph();

                                        FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                                        p.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                        //Set data format
                                        TextRange TR;
                                        TR = p.AppendText(value_vi);

                                        //TR.CharacterFormat.FontName = "Arial";

                                        TR.CharacterFormat.FontSize = 10;

                                        //TR.CharacterFormat.Bold = true;
                                        if (value_en != null)
                                        {
                                            p = FRow.Cells[i].AddParagraph();
                                            p.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                            //Set data format
                                            TR = p.AppendText(value_en);

                                            //TR.CharacterFormat.FontName = "Arial";

                                            TR.CharacterFormat.FontSize = 10;
                                            TR.CharacterFormat.Italic = true;


                                        }

                                    }
                                    else if (column.target.value_type == "varchar")
                                    {
                                        Paragraph p = FRow.Cells[i].AddParagraph();

                                        FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                                        p.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                        //Set data format
                                        TextRange TR;
                                        TR = p.AppendText(data_point.value_text);

                                        //TR.CharacterFormat.FontName = "Arial";

                                        TR.CharacterFormat.FontSize = 10;

                                        //TR.CharacterFormat.Bold = true;
                                    }
                                    else
                                    {
                                        Paragraph p = FRow.Cells[i].AddParagraph();

                                        FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                                        p.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                        //Set data format
                                        TextRange TR;
                                        TR = p.AppendText(data_point.value.Value.ToString("#,##0.##"));

                                        //TR.CharacterFormat.FontName = "Arial";

                                        TR.CharacterFormat.FontSize = 10;

                                        //TR.CharacterFormat.Bold = true;
                                    }


                                }
                                else
                                {
                                    FRow.Cells[i].CellFormat.BackColor = Color.LightGray;
                                }
                                i++;
                            }

                            row_current++;
                        }
                    }

                    if ("Max" == "Max")
                    {
                        //// Row Max
                        //Set the first row as table header
                        FRow = table1.Rows[row_current];

                        //FRow.IsHeader = true;
                        //Set the height and color of the first row
                        FRow.Height = 30;
                        i = 0;

                        //table1.ApplyHorizontalMerge(row_current, 0, 1);

                        p1 = FRow.Cells[i].AddParagraph();
                        FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                        p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                        TR1 = p1.AppendText("Max");

                        //TR1.CharacterFormat.FontName = "Arial";

                        TR1.CharacterFormat.FontSize = 10;

                        TR1.CharacterFormat.Bold = true;
                        i++;


                        foreach (var column in batch)
                        {
                            //Set alignment for cells
                            var data_point = point.data.Where(d => d.target_id == column.id).Select(d => d.value).Max();
                            if (data_point != null && column.target.value_type == "float")
                            {
                                Paragraph p = FRow.Cells[i].AddParagraph();

                                FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                                p.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                //Set data format
                                TextRange TR = p.AppendText(data_point.Value.ToString("#,##0.##"));

                                //TR.CharacterFormat.FontName = "Arial";

                                TR.CharacterFormat.FontSize = 10;

                                //TR.CharacterFormat.Bold = true;

                            }
                            else
                            {
                                FRow.Cells[i].CellFormat.BackColor = Color.LightGray;
                            }
                            i++;
                        }
                        row_current++;

                    }
                    if ("Min" == "Min")
                    {
                        //// Row Min
                        //Set the first row as table header
                        FRow = table1.Rows[row_current];

                        //FRow.IsHeader = true;
                        //Set the height and color of the first row
                        FRow.Height = 30;
                        i = 0;

                        //table1.ApplyHorizontalMerge(row_current, 0, 1);

                        p1 = FRow.Cells[i].AddParagraph();
                        FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                        p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                        TR1 = p1.AppendText("Min");

                        //TR1.CharacterFormat.FontName = "Arial";

                        TR1.CharacterFormat.FontSize = 10;

                        TR1.CharacterFormat.Bold = true;
                        i++;



                        foreach (var column in batch)
                        {
                            //Set alignment for cells
                            var data_point = point.data.Where(d => d.target_id == column.id).Select(d => d.value).Min();
                            if (data_point != null && column.target.value_type == "float")
                            {
                                Paragraph p = FRow.Cells[i].AddParagraph();

                                FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                                p.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                //Set data format
                                TextRange TR = p.AppendText(data_point.Value.ToString("#,##0.##"));

                                //TR.CharacterFormat.FontName = "Arial";

                                TR.CharacterFormat.FontSize = 10;

                                //TR.CharacterFormat.Bold = true;

                            }
                            else
                            {
                                FRow.Cells[i].CellFormat.BackColor = Color.LightGray;
                            }
                            i++;
                        }
                        row_current++;

                    }

                    if (1 == 1)
                    {
                        //// Row Kết quả trước đó
                        //Set the first row as table header
                        FRow = table1.Rows[row_current];

                        //FRow.IsHeader = true;
                        //Set the height and color of the first row
                        FRow.Height = 30;
                        i = 0;

                        table1.ApplyHorizontalMerge(row_current, 0, batch.Count());

                        p1 = FRow.Cells[i].AddParagraph();
                        FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                        p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                        TR1 = p1.AppendText("Kết quả trước đó / ");

                        //TR1.CharacterFormat.FontName = "Arial";

                        TR1.CharacterFormat.FontSize = 10;

                        TR1.CharacterFormat.Bold = true;

                        //p1 = FRow.Cells[i].AddParagraph();
                        //p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                        TR1 = p1.AppendText("Results of previous");

                        //TR1.CharacterFormat.FontName = "Arial";

                        TR1.CharacterFormat.FontSize = 10;

                        TR1.CharacterFormat.Italic = true;
                        i++;
                        row_current++;
                    }

                    if ("Max_prev" == "Max_prev")
                    {
                        //// Row Max Trước đó
                        //Set the first row as table header
                        FRow = table1.Rows[row_current];

                        //FRow.IsHeader = true;
                        //Set the height and color of the first row
                        FRow.Height = 30;
                        i = 0;

                        //table1.ApplyHorizontalMerge(row_current, 0, 1);

                        p1 = FRow.Cells[i].AddParagraph();
                        FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                        p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                        TR1 = p1.AppendText("Max");

                        //TR1.CharacterFormat.FontName = "Arial";

                        TR1.CharacterFormat.FontSize = 10;

                        TR1.CharacterFormat.Bold = true;
                        i++;



                        foreach (var column in batch)
                        {
                            //Set alignment for cells
                            var data_point = list_prev.Where(d => d.target_id == column.id && d.point_id == point.id).Select(d => d.value).Max();
                            if (data_point != null && column.target.value_type == "float")
                            {
                                Paragraph p = FRow.Cells[i].AddParagraph();

                                FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                                p.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                //Set data format
                                TextRange TR = p.AppendText(data_point.Value.ToString("#,##0.##"));

                                //TR.CharacterFormat.FontName = "Arial";

                                TR.CharacterFormat.FontSize = 10;

                                //TR.CharacterFormat.Bold = true;

                            }
                            else
                            {
                                FRow.Cells[i].CellFormat.BackColor = Color.LightGray;
                            }
                            i++;
                        }
                        row_current++;

                    }
                    if ("Min_prev" == "Min_prev")
                    {
                        //// Row Min Trước đó
                        //Set the first row as table header
                        FRow = table1.Rows[row_current];

                        //FRow.IsHeader = true;
                        //Set the height and color of the first row
                        FRow.Height = 30;
                        i = 0;

                        //table1.ApplyHorizontalMerge(row_current, 0, 1);

                        p1 = FRow.Cells[i].AddParagraph();
                        FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                        p1.Format.HorizontalAlignment = HorizontalAlignment.Center;
                        TR1 = p1.AppendText("Min");

                        //TR1.CharacterFormat.FontName = "Arial";

                        TR1.CharacterFormat.FontSize = 10;

                        TR1.CharacterFormat.Bold = true;
                        i++;




                        foreach (var column in batch)
                        {
                            //Set alignment for cells
                            var data_point = list_prev.Where(d => d.target_id == column.id && d.point_id == point.id).Select(d => d.value).Min();
                            if (data_point != null && column.target.value_type == "float")
                            {
                                Paragraph p = FRow.Cells[i].AddParagraph();

                                FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                                p.Format.HorizontalAlignment = HorizontalAlignment.Center;
                                //Set data format
                                TextRange TR = p.AppendText(data_point.Value.ToString("#,##0.##"));

                                //TR.CharacterFormat.FontName = "Arial";

                                TR.CharacterFormat.FontSize = 10;

                                //TR.CharacterFormat.Bold = true;

                            }
                            else
                            {
                                FRow.Cells[i].CellFormat.BackColor = Color.LightGray;
                            }
                            i++;
                        }
                        row_current++;
                    }




                    tables.Add(new Label1()
                    {
                        key = point.id,
                        table_ketqua = text
                    });
                }

                //Trend
                var groups_trend = results.Where(d => d.target.value_type == "float")
                               .GroupBy(e => new { e.point.frequency_id, e.point.frequency, e.point.frequency_en }).Select(e => new Label1()
                               {
                                   key = e.Key.frequency_id.Value,
                                   label = e.Key.frequency,
                                   label_en = e.Key.frequency_en,
                                   children = e
                                   .GroupBy(f => f.target).Select(f => new Label1()
                                   {
                                       key = f.Key.id,
                                       label = f.Key.name,
                                       label_en = f.Key.name_en,
                                       parent = e.Key.frequency_id.Value,
                                   }).ToList()
                               }).ToList();
                var fres = new List<Label1>();
                var targets = new List<Label1>();

                foreach (var f in groups_trend)
                {
                    fres.Add(f);
                    foreach (var t in f.children)
                    {
                        t.image_ketqua = $"chart_{f.key}_{t.key}";
                        t.image_link_ketqua = $"link_{f.key}_{t.key}";
                        var chartData = _context.ChartDataModel.Where(d => d.timestamp == timestamp && d.key == t.image_ketqua).FirstOrDefault();
                        if (chartData != null)
                        {
                            t.link_ketqua = _configuration["Application:Trend:link"] + "report/chart/" + Uri.EscapeDataString(_AesOperation.EncryptString(chartData.id.ToString()));
                        }
                        targets.Add(t);
                    }
                }

                List<DictionaryEntry> relationsList = new List<DictionaryEntry>();
                relationsList.Add(new DictionaryEntry("points", string.Empty));
                relationsList.Add(new DictionaryEntry("children", "key = %points.id%"));

                MailMergeDataSet mailMergeDataSet = new MailMergeDataSet();
                mailMergeDataSet.Add(new MailMergeDataTable("points", points));
                mailMergeDataSet.Add(new MailMergeDataTable("children", tables));
                document.MailMerge.ExecuteWidthNestedRegion(mailMergeDataSet, relationsList);


                foreach (KeyValuePair<string, Spire.Doc.Table> entry in replacements_table)
                {
                    TextSelection selection = document.FindString(entry.Key, true, true);
                    if (selection == null)
                    {
                        var table = entry.Value;
                        //table
                        continue;
                    }
                    TextRange range = selection.GetAsOneRange();
                    Paragraph paragraph = range.OwnerParagraph;
                    Body body = paragraph.OwnerTextBody;
                    int index = body.ChildObjects.IndexOf(paragraph);
                    //var fontsize = paragrap

                    body.ChildObjects.Remove(paragraph);
                    body.ChildObjects.Insert(index, entry.Value);
                }

                relationsList = new List<DictionaryEntry>();
                //relationsList.Add(new DictionaryEntry("groups", string.Empty));
                //relationsList.Add(new DictionaryEntry("children", "key = %groups.key%"));

                relationsList.Add(new DictionaryEntry("fres", string.Empty));
                relationsList.Add(new DictionaryEntry("targets", "parent = %fres.key%"));

                mailMergeDataSet = new MailMergeDataSet();
                //mailMergeDataSet.Add(new MailMergeDataTable("groups", groups));
                //mailMergeDataSet.Add(new MailMergeDataTable("children", tables));

                mailMergeDataSet.Add(new MailMergeDataTable("fres", fres));
                mailMergeDataSet.Add(new MailMergeDataTable("targets", targets));

                document.MailMerge.ExecuteWidthNestedRegion(mailMergeDataSet, relationsList);
                //Set Font Style and Size

                ParagraphStyle style = new ParagraphStyle(document);

                style.Name = "FontStyle";
                style.CharacterFormat.FontSize = 12;

                document.Styles.Add(style);

                foreach (var target in targets)
                {
                    TextSelection selection = document.FindString(target.image_ketqua, true, true);
                    TextSelection selection1 = document.FindString(target.image_link_ketqua, true, true);
                    if (selection == null)
                    {
                        //var table = entry.Value;
                        //table
                        continue;
                    }
                    Image image = Image.FromFile($"./wwwroot/tmp/{timestamp}/{target.image_ketqua}.png");
                    Paragraph paragraph = section.AddParagraph();

                    DocPicture pic = new DocPicture(document);
                    pic.LoadImage(image);
                    pic.Width = 750;
                    pic.Height = 350;
                    paragraph.AppendHyperlink(target.link_ketqua, "Xem biểu đồ", HyperlinkType.WebLink);
                    paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center;

                    paragraph.ApplyStyle(style.Name);

                    var range = selection.GetAsOneRange();
                    var index = range.OwnerParagraph.ChildObjects.IndexOf(range);
                    range.OwnerParagraph.ChildObjects.Insert(index, pic);
                    range.OwnerParagraph.ChildObjects.Remove(range);

                    TextRange range1 = selection1.GetAsOneRange();
                    Paragraph paragraph1 = range1.OwnerParagraph;
                    Body body = paragraph1.OwnerTextBody;
                    int index1 = body.ChildObjects.IndexOf(paragraph1);
                    //var fontsize = paragrap

                    body.ChildObjects.Remove(paragraph1);
                    body.ChildObjects.Insert(index1, paragraph);
                    //var range1 = selection1.GetAsOneRange();
                    //var index1 = range1.OwnerParagraph.ChildObjects.IndexOf(range1);
                    //range1.OwnerParagraph.ChildObjects.Insert(index1, paragraph);
                    //range1.OwnerParagraph.ChildObjects.Remove(range1);

                }


                document.SaveToFile("./wwwroot" + documentPath, Spire.Doc.FileFormat.Docx);


                return Json(new { success = true, link = Domain + documentPath });
            }

            return Json(new { success = false });

        }
        [HttpPost]
        public async Task<JsonResult> drawChart(LimitModel LimitModel, List<int> list_point, int type_bc, DateTime? date_from_prev, DateTime? date_to_prev)
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
            if (date_from_prev != null && date_from_prev.Value.Kind == DateTimeKind.Utc)
            {
                date_from_prev = date_from_prev.Value.ToLocalTime();
            }
            if (date_to_prev != null && date_to_prev.Value.Kind == DateTimeKind.Utc)
            {
                date_to_prev = date_to_prev.Value.ToLocalTime();
            }
            if (LimitModel.object_id == 2)
            {
                var from = date_from_prev != null ? date_from_prev : date_from;
                var results = _context.ResultModel.Where(d => d.deleted_at == null && list_point.Contains(d.point_id.Value) && d.date >= from && d.date <= date_to)
                    .Include(d => d.target)
                    .Include(d => d.point)
                    .ThenInclude(d => d.location)
                    .Where(h => h.target.value_type == "float").ToList();

                var groups = results
                                .GroupBy(e => new { e.point.frequency_id, e.point.frequency }).Select(e => new Label()
                                {
                                    key = e.Key.frequency_id.Value,
                                    label = e.Key.frequency,
                                    type = "frequency",
                                    timestamp = new DateTimeOffset(DateTime.UtcNow).ToUnixTimeSeconds(),
                                    children = e
                                    .GroupBy(f => f.target).Select(f => new Label()
                                    {
                                        key = f.Key.id,
                                        label = f.Key.name,
                                        type = "target",
                                        timestamp = new DateTimeOffset(DateTime.UtcNow).ToUnixTimeSeconds(),
                                        children = f.GroupBy(d => d.point.location).Select(d => new Label()
                                        {
                                            key = d.Key.id,
                                            label = d.Key.name,
                                            type = "location",
                                            timestamp = new DateTimeOffset(DateTime.UtcNow).ToUnixTimeSeconds(),
                                            data = d.ToList()
                                        }).ToList(),
                                        data = f.ToList(),
                                    }).ToList()
                                }).ToList();
                var list_pointStyle = new List<string>()
                {
                    "circle",
                    "cross",
                    "crossRot",
                    "dash",
                    "line",
                    "rect",
                    "rectRounded",
                    "rectRot",
                    "star",
                    "triangle"
                };

                foreach (var frequency in groups)
                {
                    foreach (var target in frequency.children)
                    {
                        foreach (var location in target.children)
                        {
                            //target.chart = target.data.Select(d => d.point.code).ToList();
                            //var limit_all = _context.LimitModel.Where(d => d.target_id == target.key && d.date_effect <= denngay && d.deleted_at == null).Include(d => d.points).ToList();
                            var data = location.data;
                            var list_point1 = data.Select(d => d.point_id).Distinct().ToList();
                            //limit_all = limit_all.Where(d => d.list_point.Intersect(list_point1).Any()).ToList();
                            var labels = data.GroupBy(g => g.date).Select(g => g.Key.Value).OrderBy(g => g).ToList();
                            if (date_to_prev != null)
                            {
                                labels.Add(date_to_prev.Value);
                                labels = labels.OrderBy(g => g).ToList();
                            }
                            var datasets = data.GroupBy(g => g.point).Select(g => new Dataset()
                            {
                                label = g.Key.code,
                                type = "line",
                                borderWidth = 1,
                                spanGaps = true,
                                borderColor = g.Key.color,
                                backgroundColor = g.Key.color,
                                results = g.ToList(),
                                data = new List<decimal?>()
                            }).ToList();
                            var annotations = new Dictionary<string, Annotations>();
                            var name_line_ghcb = "";
                            var name_line_ghhd = "";
                            var stt_action = 0;
                            var stt_alert = 0;
                            decimal? limit_action_prev = 0;
                            decimal? limit_alert_prev = 0;
                            var row = 0;
                            var max = data.Select(g => g.value).Max();
                            var suggestedMax = max + (max * 50 / 100);
                            foreach (var label in labels)
                            {

                                var stt = 0;
                                foreach (var d in datasets)
                                {
                                    d.pointStyle = list_pointStyle[stt];
                                    var finddata = d.results.Where(d => d.date == label).FirstOrDefault();
                                    //var point_id = finddata.point_id;

                                    var value = finddata != null ? finddata.value : null;
                                    d.data.Add(value);
                                    //var date = d.date.Value.ToString("yyyy-MM-dd");
                                    //var point_code = d.point.code;
                                    stt = list_pointStyle.Count - 1 > stt ? stt + 1 : 0;
                                }
                                var limit_action = data.Where(d => d.date == label).Select(g => g.limit_action).Min();

                                if (limit_action != null)
                                {
                                    if (limit_action == limit_action_prev)
                                    {
                                        annotations[name_line_ghhd].xMax++;
                                    }
                                    else
                                    {
                                        limit_action_prev = limit_action;
                                        name_line_ghhd = "line_ghhd_" + stt_action;
                                        annotations.Add(name_line_ghhd, new Annotations()
                                        {
                                            xMin = row,
                                            xMax = row + 1,
                                            yMin = limit_action.Value,
                                            yMax = limit_action.Value,
                                            type = "line",
                                            borderWidth = 3,
                                            borderColor = "red"
                                        });
                                        annotations.Add("callout_ghhd_" + stt_action, new Annotations()
                                        {
                                            type = "label",
                                            xValue = row,
                                            yValue = limit_action.Value,
                                            xAdjust = 50,
                                            yAdjust = -20,
                                            content = new List<string>() { "Action Limit", limit_action.Value.ToString("#,##0.##") },
                                            callout = new Callout()
                                            {
                                                display = true,
                                                side = 5
                                            }
                                        });

                                        ///
                                        if (limit_action + (limit_action * 50 / 100) > suggestedMax)
                                        {
                                            suggestedMax = limit_action + (limit_action * 50 / 100);
                                        }
                                        stt_action++;
                                    }
                                }
                                var limit_alert = data.Where(d => d.date == label).Select(g => g.limit_alert).Min();

                                if (limit_alert != null)
                                {
                                    if (limit_alert == limit_alert_prev)
                                    {
                                        annotations[name_line_ghcb].xMax++;
                                    }
                                    else
                                    {
                                        limit_alert_prev = limit_alert;
                                        name_line_ghcb = "line_ghcb_" + stt_alert;
                                        annotations.Add(name_line_ghcb, new Annotations()
                                        {
                                            xMin = row,
                                            xMax = row + 1,
                                            yMin = limit_alert.Value,
                                            yMax = limit_alert.Value,
                                            type = "line",
                                            borderWidth = 3,
                                            borderColor = "orange"
                                        });

                                        annotations.Add("callout_ghcb_" + stt_alert, new Annotations()
                                        {
                                            type = "label",
                                            xValue = row,
                                            yValue = limit_alert.Value,
                                            xAdjust = 50,
                                            yAdjust = -20,
                                            content = new List<string>() { "Alert Limit", limit_alert.Value.ToString("#,##0.##") },
                                            callout = new Callout()
                                            {
                                                display = true,
                                                side = 5
                                            }
                                        });
                                        ///
                                        if (limit_action + (limit_action * 50 / 100) > suggestedMax)
                                        {
                                            suggestedMax = limit_action + (limit_action * 50 / 100);
                                        }
                                        stt_alert++;
                                    }
                                }

                                if (label == date_to_prev)
                                {
                                    annotations.Add("grandline", new Annotations()
                                    {
                                        xMin = row,
                                        xMax = row,
                                        yMin = 0,
                                        yMax = suggestedMax.Value,
                                        type = "line",
                                        borderWidth = 5,
                                        borderColor = "gray"
                                    });
                                    var yValue = suggestedMax - (suggestedMax * 10 / 100);
                                    annotations.Add("label_prev", new Annotations()
                                    {
                                        type = "label",
                                        xValue = row / 2,
                                        yValue = yValue.Value,
                                        borderColor = "gray",
                                        borderWidth = 1,
                                        font = new Font()
                                        {
                                            weight = "bold"
                                        },
                                        content = new List<string>() { "Kết quả phân tích xu hướng lần trước", "Results of previous trend analysis" },
                                    });
                                    annotations.Add("label_current", new Annotations()
                                    {
                                        type = "label",
                                        xValue = (row + labels.Count()) / 2,
                                        yValue = yValue.Value,
                                        borderColor = "gray",
                                        borderWidth = 1,
                                        font = new Font()
                                        {
                                            weight = "bold"
                                        },
                                        content = new List<string>() { "Kết quả phân tích xu hướng hiện tại", "Results of current trend analysis" },
                                    });
                                }
                                row++;
                            }
                            //if (ghcb.data.Count > 0)
                            //{
                            //    datasets.Add(ghcb);
                            //    datasets.Add(ghhd);
                            //}
                            location.chart = new Chart1()
                            {
                                labels = labels.Select(d => d.ToString("yyyy-MM-dd")).ToList(),
                                datasets = datasets,
                                yTitle = data.FirstOrDefault().target.unit,
                                suggestedMax = suggestedMax,
                                annotations = annotations
                            };
                        }
                    }
                }
                return Json(new { success = true, data = groups }, new System.Text.Json.JsonSerializerOptions()
                {
                    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
                });
            }
            else
            {
                var from = date_from_prev != null ? date_from_prev : date_from;
                var results = _context.ResultModel.Where(d => d.deleted_at == null && list_point.Contains(d.point_id.Value) && d.date >= from && d.date <= date_to)
                    .Include(d => d.target)
                    .Include(d => d.point)
                    .ThenInclude(d => d.location)
                    .Where(h => h.target.value_type == "float").ToList();

                var groups = results
                                .GroupBy(e => new { e.point.frequency_id, e.point.frequency }).Select(e => new Label()
                                {
                                    key = e.Key.frequency_id.Value,
                                    label = e.Key.frequency,
                                    type = "frequency",
                                    timestamp = new DateTimeOffset(DateTime.UtcNow).ToUnixTimeSeconds(),
                                    children = e
                                    .GroupBy(f => f.target).Select(f => new Label()
                                    {
                                        key = f.Key.id,
                                        label = f.Key.name,
                                        type = "target",
                                        timestamp = new DateTimeOffset(DateTime.UtcNow).ToUnixTimeSeconds(),
                                        data = f.ToList(),
                                    }).ToList()
                                }).ToList();
                var list_pointStyle = new List<string>()
                {
                    "circle",
                    "cross",
                    "crossRot",
                    "dash",
                    "line",
                    "rect",
                    "rectRounded",
                    "rectRot",
                    "star",
                    "triangle"
                };

                foreach (var frequency in groups)
                {
                    foreach (var target in frequency.children)
                    {

                        //target.chart = target.data.Select(d => d.point.code).ToList();
                        //var limit_all = _context.LimitModel.Where(d => d.target_id == target.key && d.date_effect <= denngay && d.deleted_at == null).Include(d => d.points).ToList();
                        var data = target.data;
                        var list_point1 = data.Select(d => d.point_id).Distinct().ToList();
                        //limit_all = limit_all.Where(d => d.list_point.Intersect(list_point1).Any()).ToList();
                        var labels = data.GroupBy(g => g.date).Select(g => g.Key.Value).OrderBy(g => g).ToList();
                        if (date_to_prev != null)
                        {
                            labels.Add(date_to_prev.Value);
                            labels = labels.OrderBy(g => g).Distinct().ToList();
                        }
                        var datasets = data.GroupBy(g => g.point).Select(g => new Dataset()
                        {
                            label = g.Key.code,
                            type = "line",
                            borderWidth = 1,
                            spanGaps = true,
                            borderColor = g.Key.color,
                            backgroundColor = g.Key.color,
                            results = g.ToList(),
                            data = new List<decimal?>()
                        }).ToList();
                        var annotations = new Dictionary<string, Annotations>();
                        var name_line_ghcb = "";
                        var name_line_ghhd = "";
                        var stt_action = 0;
                        var stt_alert = 0;
                        decimal? limit_action_prev = 0;
                        decimal? limit_alert_prev = 0;
                        var row = 0;
                        var max = data.Select(g => g.value).Max();
                        var suggestedMax = max + (max * 50 / 100);
                        foreach (var label in labels)
                        {

                            var stt = 0;
                            foreach (var d in datasets)
                            {
                                d.pointStyle = list_pointStyle[stt];
                                var finddata = d.results.Where(d => d.date == label).FirstOrDefault();
                                //var point_id = finddata.point_id;

                                var value = finddata != null ? finddata.value : null;
                                d.data.Add(value);
                                //var date = d.date.Value.ToString("yyyy-MM-dd");
                                //var point_code = d.point.code;
                                stt = list_pointStyle.Count - 1 > stt ? stt + 1 : 0;
                            }
                            var limit_action = data.Where(d => d.date == label).Select(g => g.limit_action).Min();

                            if (limit_action != null)
                            {
                                if (limit_action == limit_action_prev)
                                {
                                    annotations[name_line_ghhd].xMax++;
                                }
                                else
                                {
                                    limit_action_prev = limit_action;
                                    name_line_ghhd = "line_ghhd_" + stt_action;
                                    annotations.Add(name_line_ghhd, new Annotations()
                                    {
                                        xMin = row,
                                        xMax = row + 1,
                                        yMin = limit_action.Value,
                                        yMax = limit_action.Value,
                                        type = "line",
                                        borderWidth = 3,
                                        borderColor = "red"
                                    });
                                    annotations.Add("callout_ghhd_" + stt_action, new Annotations()
                                    {
                                        type = "label",
                                        xValue = row,
                                        yValue = limit_action.Value,
                                        xAdjust = 50,
                                        yAdjust = -20,
                                        content = new List<string>() { "Action Limit", limit_action.Value.ToString("#,##0.##") },
                                        callout = new Callout()
                                        {
                                            display = true,
                                            side = 5
                                        }
                                    });

                                    ///
                                    if (limit_action + (limit_action * 50 / 100) > suggestedMax)
                                    {
                                        suggestedMax = limit_action + (limit_action * 50 / 100);
                                    }
                                    stt_action++;
                                }
                            }
                            var limit_alert = data.Where(d => d.date == label).Select(g => g.limit_alert).Min();

                            if (limit_alert != null)
                            {
                                if (limit_alert == limit_alert_prev)
                                {
                                    annotations[name_line_ghcb].xMax++;
                                }
                                else
                                {
                                    limit_alert_prev = limit_alert;
                                    name_line_ghcb = "line_ghcb_" + stt_alert;
                                    annotations.Add(name_line_ghcb, new Annotations()
                                    {
                                        xMin = row,
                                        xMax = row + 1,
                                        yMin = limit_alert.Value,
                                        yMax = limit_alert.Value,
                                        type = "line",
                                        borderWidth = 3,
                                        borderColor = "orange"
                                    });

                                    annotations.Add("callout_ghcb_" + stt_alert, new Annotations()
                                    {
                                        type = "label",
                                        xValue = row,
                                        yValue = limit_alert.Value,
                                        xAdjust = 50,
                                        yAdjust = -20,
                                        content = new List<string>() { "Alert Limit", limit_alert.Value.ToString("#,##0.##") },
                                        callout = new Callout()
                                        {
                                            display = true,
                                            side = 5
                                        }
                                    });
                                    ///
                                    if (limit_action + (limit_action * 50 / 100) > suggestedMax)
                                    {
                                        suggestedMax = limit_action + (limit_action * 50 / 100);
                                    }
                                    stt_alert++;
                                }
                            }

                            if (label == date_to_prev)
                            {
                                annotations.Add("grandline", new Annotations()
                                {
                                    xMin = row,
                                    xMax = row,
                                    yMin = 0,
                                    yMax = suggestedMax.Value,
                                    type = "line",
                                    borderWidth = 5,
                                    borderColor = "gray"
                                });
                                var yValue = suggestedMax - (suggestedMax * 10 / 100);
                                annotations.Add("label_prev", new Annotations()
                                {
                                    type = "label",
                                    xValue = row / 2,
                                    yValue = yValue.Value,
                                    borderColor = "gray",
                                    borderWidth = 1,
                                    font = new Font()
                                    {
                                        weight = "bold"
                                    },
                                    content = new List<string>() { "Kết quả phân tích xu hướng lần trước", "Results of previous trend analysis" },
                                });
                                annotations.Add("label_current", new Annotations()
                                {
                                    type = "label",
                                    xValue = (row + labels.Count()) / 2,
                                    yValue = yValue.Value,
                                    borderColor = "gray",
                                    borderWidth = 1,
                                    font = new Font()
                                    {
                                        weight = "bold"
                                    },
                                    content = new List<string>() { "Kết quả phân tích xu hướng hiện tại", "Results of current trend analysis" },
                                });
                            }
                            row++;
                        }
                        //if (ghcb.data.Count > 0)
                        //{
                        //    datasets.Add(ghcb);
                        //    datasets.Add(ghhd);
                        //}
                        target.chart = new Chart1()
                        {
                            labels = labels.Select(d => d.ToString("yyyy-MM-dd")).ToList(),
                            datasets = datasets,
                            yTitle = data.FirstOrDefault().target.unit,
                            suggestedMax = suggestedMax,
                            annotations = annotations
                        };

                    }
                }
                return Json(new { success = true, data = groups }, new System.Text.Json.JsonSerializerOptions()
                {
                    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
                });
            }
            return Json(new { success = false });
        }

        [HttpPost]
        public async Task<JsonResult> uploadImage(List<Files> files)
        {
            var timestamp = new DateTimeOffset(DateTime.UtcNow).ToUnixTimeSeconds();
            bool exists = System.IO.Directory.Exists("./wwwroot/tmp/" + timestamp);

            if (!exists)
                System.IO.Directory.CreateDirectory("./wwwroot/tmp/" + timestamp);
            foreach (var file in files)
            {
                var base64String = file.image.Replace("data:image/png;base64,", "");

                // Convert Base64 string to byte array
                byte[] imageBytes = Convert.FromBase64String(base64String);

                // Create an image from the byte array
                using (MemoryStream ms = new MemoryStream(imageBytes))
                {
                    Image image = Image.FromStream(ms);

                    // Save or use the image as required
                    image.Save("./wwwroot/tmp/" + timestamp + "/" + file.name + ".png", ImageFormat.Png);
                    var chartdata = new ChartDataModel()
                    {
                        timestamp = timestamp,
                        key = file.name,
                        data = file.chart
                    };
                    _context.Add(chartdata);
                }
            }
            _context.SaveChanges();
            return Json(new { success = true, timestamp = timestamp });
        }
        public JsonResult GetChartData(string id_text)
        {
            if (id_text == null)
                return Json(null);
            var plain_text = _AesOperation.DecryptString(id_text);
            var id = Int32.Parse(plain_text);
            var data = _context.ChartDataModel.Where(d => d.id == id).FirstOrDefault();
            return Json(data, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
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

        static string ColumnIndexToColumnLetter(int colIndex)
        {
            int div = colIndex;
            string colLetter = String.Empty;
            int mod = 0;

            while (div > 0)
            {
                mod = (div - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                div = (int)((div - mod) / 26);
            }
            return colLetter;
        }


    }
    public class Files
    {
        public string name { get; set; }
        public string image { get; set; }
        public string chart { get; set; }
    }
    class Points
    {
        public int id { get; set; }
        public string point { get; set; }

        public DateTime? sort { get; set; }
        public string min_ddd { get; set; }
        public string max_ddd { get; set; }

        public string min_toc { get; set; }
        public string max_toc { get; set; }
        public string min_tamc { get; set; }
        public string max_tamc { get; set; }

        public string min_p_ddd { get; set; }
        public string max_p_ddd { get; set; }

        public string min_p_toc { get; set; }
        public string max_p_toc { get; set; }
        public string min_p_tamc { get; set; }
        public string max_p_tamc { get; set; }
        public List<ResultModel> data { get; set; }
    }
    class Details
    {
        public int id { get; set; }
        public string date { get; set; }

        public DateTime? sort { get; set; }
        public string tinhchat { get; set; }
        public string nitrat { get; set; }

        public string ddd { get; set; }
        public string toc { get; set; }
        public string tamc { get; set; }
        public string vsgb { get; set; }
        public List<ResultModel> data { get; set; }
    }

    public class Label1
    {
        public int? key { set; get; }
        public string text { set; get; }
        public string label { set; get; }
        public string label_en { set; get; }
        public int? parent { set; get; }
        public string table_ketqua { set; get; }
        public string image_ketqua { set; get; }
        public string image_link_ketqua { set; get; }
        public string link_ketqua { set; get; }
        public long timestamp { get; set; }
        public List<Label1> children { get; set; }
        public List<string> list { get; set; }
        public Chart1? chart { get; set; }
        public List<ResultModel>? data { get; set; }



    }
}