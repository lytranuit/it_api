

using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Spire.Xls;
using System.Collections;
using System.Data;
using System.Security.Cryptography.X509Certificates;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using Vue.Data;
using Vue.Models;
using iText.Commons.Actions.Contexts;
using Microsoft.EntityFrameworkCore.Infrastructure;
using System.Diagnostics;
using Microsoft.CodeAnalysis;
using System.Drawing;
using System.Security.Cryptography;
using System.Text;

namespace it_template.Areas.Trend.Controllers
{

    [Area("Trend")]
    public class ImportController : Controller
    {
        private readonly IConfiguration _configuration;
        private readonly ItContext _context;
        public ImportController(ItContext context, IConfiguration configuration) : base()
        {
            _context = context;
            _configuration = configuration;
            var listener = _context.GetService<DiagnosticSource>();
            (listener as DiagnosticListener).SubscribeWithAdapter(new CommandInterceptor());
        }

        public async Task<IActionResult> import1T()
        {
            return Ok();
            // Khởi tạo workbook để đọc
            Spire.Xls.Workbook book = new Spire.Xls.Workbook();
            book.LoadFromFile("./wwwroot/data/trend/nuoc/Raw data of 1T.xlsx", ExcelVersion.Version2013);

            Spire.Xls.Worksheet sheet = book.Worksheets[0];
            var lastrow = sheet.LastDataRow;
            // nếu vẫn chưa gặp end thì vẫn lấy data
            Console.WriteLine(lastrow);
            var list_Result = new List<ResultModel>();
            for (int rowIndex = 4; rowIndex < lastrow; rowIndex++)
            {
                // lấy row hiện tại
                var nowRow = sheet.Rows[rowIndex];
                if (nowRow == null)
                    continue;
                // vì ta dùng 3 cột A, B, C => data của ta sẽ như sau
                //int numcount = nowRow.Cells.Count;
                //for(int y = 0;y<numcount - 1 ;y++)
                var code_vitri = nowRow.Cells[1] != null ? nowRow.Cells[1].Value : null;
                // Xuất ra thông tin lên màn hình
                Console.WriteLine("MS: {0} ", code_vitri);
                Console.WriteLine("rowIndex: {0} ", rowIndex);

                if (code_vitri == null)
                    continue;

                var tansuat = nowRow.Cells[0] != null ? nowRow.Cells[0].Value : null;
                DateTime? date = nowRow.Cells[2] != null ? nowRow.Cells[2].DateTimeValue : null;
                var tansuat_id = 1;
                switch (tansuat)
                {
                    case "Daily":
                        tansuat_id = 1;
                        break;
                    case "Twice per month":
                        tansuat_id = 2;
                        break;

                }
                var chitieu_9 = nowRow.Cells[3] != null && nowRow.Cells[3].Value != "NA" && nowRow.Cells[3].Value != "" ? nowRow.Cells[3].Value : null;
                var chitieu_10 = nowRow.Cells[4] != null && nowRow.Cells[4].Value != "NA" && nowRow.Cells[4].Value != "" ? nowRow.Cells[4].Value : null;
                decimal? chitieu_11 = nowRow.Cells[5] != null && nowRow.Cells[5].Value != "NA" && nowRow.Cells[5].Value != "" ? (decimal)nowRow.Cells[5].NumberValue : null;
                decimal? chitieu_12 = nowRow.Cells[6] != null && nowRow.Cells[6].Value != "NA" && nowRow.Cells[6].Value != "" ? (decimal)nowRow.Cells[6].NumberValue : null;
                decimal? chitieu_13 = nowRow.Cells[7] != null && nowRow.Cells[7].Value != "NA" && nowRow.Cells[7].Value != "" ? (decimal)nowRow.Cells[7].NumberValue : null;
                var chitieu_14 = nowRow.Cells[8] != null && nowRow.Cells[8].Value != "NA" && nowRow.Cells[8].Value != "" ? nowRow.Cells[8].Value : null;

                var findPoint = _context.PointModel.Where(d => d.code == code_vitri).FirstOrDefault();
                if (findPoint == null)
                {
                    findPoint = new PointModel()
                    {
                        color = ColorTranslator.ToHtml(GenerateRandomColor(code_vitri)),
                        code = code_vitri,
                        frequency_id = tansuat_id,
                        object_id = 3,
                        location_id = 7,
                        created_at = DateTime.Now,
                    };
                    _context.Add(findPoint);
                    _context.SaveChanges();
                }


                if (chitieu_9 != null)
                {
                    var chitieu_9_int = chitieu_9 == "ĐẠT" ? 1 : 0;
                    var find_result = _context.ResultModel.Where(d => d.target_id == 9 && d.date == date && d.point_id == findPoint.id).FirstOrDefault();
                    if (find_result == null)
                    {
                        var result_9 = new ResultModel()
                        {
                            point_id = findPoint.id,
                            value = chitieu_9_int,
                            target_id = 9,
                            date = date,
                            object_id = 3,
                            created_at = DateTime.Now
                        };
                        list_Result.Add(result_9);
                        _context.Add(result_9);
                    }
                    else
                    {
                        find_result.value = chitieu_9_int;
                        _context.Update(find_result);
                        //_context.SaveChanges();
                    }

                    //_context.SaveChanges();
                }
                if (chitieu_10 != null)
                {
                    var chitieu_10_int = chitieu_10 == "ĐẠT" ? 1 : 0;
                    var find_result = _context.ResultModel.Where(d => d.target_id == 10 && d.date == date && d.point_id == findPoint.id).FirstOrDefault();
                    if (find_result == null)
                    {
                        var result_10 = new ResultModel()
                        {
                            point_id = findPoint.id,
                            value = chitieu_10_int,
                            target_id = 10,
                            date = date,
                            created_at = DateTime.Now
                        };
                        list_Result.Add(result_10);
                        _context.Add(result_10);
                    }
                    else
                    {
                        find_result.value = chitieu_10_int;
                        _context.Update(find_result);
                        //_context.SaveChanges();
                    }
                    //_context.SaveChanges();
                }

                if (chitieu_11 != null)
                {
                    var find_result = _context.ResultModel.Where(d => d.target_id == 11 && d.date == date && d.point_id == findPoint.id).FirstOrDefault();
                    if (find_result == null)
                    {
                        var result_11 = new ResultModel()
                        {
                            point_id = findPoint.id,
                            value = chitieu_11,
                            target_id = 11,
                            date = date,
                            created_at = DateTime.Now
                        };
                        list_Result.Add(result_11);
                        _context.Add(result_11);
                    }
                    else
                    {
                        find_result.value = chitieu_11;
                        _context.Update(find_result);
                        //_context.SaveChanges();
                    }
                    //_context.SaveChanges();
                }
                if (chitieu_12 != null)
                {
                    var find_result = _context.ResultModel.Where(d => d.target_id == 12 && d.date == date && d.point_id == findPoint.id).FirstOrDefault();
                    if (find_result == null)
                    {
                        var result_12 = new ResultModel()
                        {
                            point_id = findPoint.id,
                            value = chitieu_12,
                            target_id = 12,
                            date = date,
                            created_at = DateTime.Now
                        };
                        list_Result.Add(result_12);
                        _context.Add(result_12);
                    }
                    else
                    {
                        find_result.value = chitieu_12;
                        _context.Update(find_result);
                        //_context.SaveChanges();
                    }
                    //_context.SaveChanges();
                }

                if (chitieu_13 != null)
                {
                    var find_result = _context.ResultModel.Where(d => d.target_id == 13 && d.date == date && d.point_id == findPoint.id).FirstOrDefault();
                    if (find_result == null)
                    {
                        var result_13 = new ResultModel()
                        {
                            point_id = findPoint.id,
                            value = chitieu_13,
                            target_id = 13,
                            date = date,
                            created_at = DateTime.Now
                        };
                        list_Result.Add(result_13);
                        _context.Add(result_13);
                    }
                    else
                    {
                        find_result.value = chitieu_13;
                        _context.Update(find_result);
                        //_context.SaveChanges();
                    }
                    //_context.SaveChanges();
                }
                //EquipmentModel EquipmentModel = new EquipmentModel { code = code, name = name_vn, name_en = name_en, created_at = DateTime.Now };
                //_context.Add(EquipmentModel);
                //_context.SaveChanges();
            }
            //_context.AddRange(list_Result);
            _context.SaveChanges();

            return Ok(list_Result);
        }


        public async Task<IActionResult> import4T()
        {
            return Ok();
            // Khởi tạo workbook để đọc
            Spire.Xls.Workbook book = new Spire.Xls.Workbook();
            book.LoadFromFile("./wwwroot/data/trend/nuoc/Raw data of 4T.xlsx", ExcelVersion.Version2013);

            Spire.Xls.Worksheet sheet = book.Worksheets[0];
            var lastrow = sheet.LastDataRow;
            // nếu vẫn chưa gặp end thì vẫn lấy data
            Console.WriteLine(lastrow);
            var list_Result = new List<ResultModel>();
            for (int rowIndex = 4; rowIndex < lastrow; rowIndex++)
            {
                // lấy row hiện tại
                var nowRow = sheet.Rows[rowIndex];
                if (nowRow == null)
                    continue;
                // vì ta dùng 3 cột A, B, C => data của ta sẽ như sau
                //int numcount = nowRow.Cells.Count;
                //for(int y = 0;y<numcount - 1 ;y++)
                var code_vitri = nowRow.Cells[1] != null ? nowRow.Cells[1].Value : null;
                // Xuất ra thông tin lên màn hình
                Console.WriteLine("MS: {0} ", code_vitri);
                Console.WriteLine("rowIndex: {0} ", rowIndex);

                if (code_vitri == null)
                    continue;

                var tansuat = nowRow.Cells[0] != null ? nowRow.Cells[0].Value : null;
                DateTime? date = nowRow.Cells[2] != null ? nowRow.Cells[2].DateTimeValue : null;
                var tansuat_id = 1;
                switch (tansuat)
                {
                    case "Daily":
                        tansuat_id = 1;
                        break;
                    case "Twice per month":
                        tansuat_id = 2;
                        break;

                }
                var chitieu_9 = nowRow.Cells[3] != null && nowRow.Cells[3].Value != "NA" && nowRow.Cells[3].Value != "" ? nowRow.Cells[3].Value : null;
                var chitieu_10 = nowRow.Cells[4] != null && nowRow.Cells[4].Value != "NA" && nowRow.Cells[4].Value != "" ? nowRow.Cells[4].Value : null;
                decimal? chitieu_11 = nowRow.Cells[5] != null && nowRow.Cells[5].Value != "NA" && nowRow.Cells[5].Value != "" ? (decimal)nowRow.Cells[5].NumberValue : null;
                decimal? chitieu_12 = nowRow.Cells[6] != null && nowRow.Cells[6].Value != "NA" && nowRow.Cells[6].Value != "" ? (decimal)nowRow.Cells[6].NumberValue : null;
                decimal? chitieu_13 = nowRow.Cells[7] != null && nowRow.Cells[7].Value != "NA" && nowRow.Cells[7].Value != "" ? (decimal)nowRow.Cells[7].NumberValue : null;
                var chitieu_14 = nowRow.Cells[8] != null && nowRow.Cells[8].Value != "NA" && nowRow.Cells[8].Value != "" ? nowRow.Cells[8].Value : null;

                var findPoint = _context.PointModel.Where(d => d.code == code_vitri).FirstOrDefault();
                if (findPoint == null)
                {
                    findPoint = new PointModel()
                    {
                        color = ColorTranslator.ToHtml(GenerateRandomColor(code_vitri)),
                        code = code_vitri,
                        frequency_id = tansuat_id,
                        object_id = 3,
                        location_id = 7,
                        created_at = DateTime.Now,
                    };
                    _context.Add(findPoint);
                    _context.SaveChanges();
                }


                if (chitieu_9 != null)
                {
                    var chitieu_9_int = chitieu_9 == "ĐẠT" ? 1 : 0;
                    var find_result = _context.ResultModel.Where(d => d.target_id == 9 && d.date == date && d.point_id == findPoint.id).FirstOrDefault();
                    if (find_result == null)
                    {
                        var result_9 = new ResultModel()
                        {
                            point_id = findPoint.id,
                            value = chitieu_9_int,
                            target_id = 9,
                            date = date,
                            object_id = 3,
                            created_at = DateTime.Now
                        };
                        list_Result.Add(result_9);
                        _context.Add(result_9);
                    }
                    else
                    {
                        find_result.value = chitieu_9_int;
                        _context.Update(find_result);
                        //_context.SaveChanges();
                    }

                    //_context.SaveChanges();
                }
                if (chitieu_10 != null)
                {
                    var chitieu_10_int = chitieu_10 == "ĐẠT" ? 1 : 0;
                    var find_result = _context.ResultModel.Where(d => d.target_id == 10 && d.date == date && d.point_id == findPoint.id).FirstOrDefault();
                    if (find_result == null)
                    {
                        var result_10 = new ResultModel()
                        {
                            point_id = findPoint.id,
                            value = chitieu_10_int,
                            target_id = 10,
                            date = date,
                            created_at = DateTime.Now
                        };
                        list_Result.Add(result_10);
                        _context.Add(result_10);
                    }
                    else
                    {
                        find_result.value = chitieu_10_int;
                        _context.Update(find_result);
                        //_context.SaveChanges();
                    }
                    //_context.SaveChanges();
                }

                if (chitieu_11 != null)
                {
                    var find_result = _context.ResultModel.Where(d => d.target_id == 11 && d.date == date && d.point_id == findPoint.id).FirstOrDefault();
                    if (find_result == null)
                    {
                        var result_11 = new ResultModel()
                        {
                            point_id = findPoint.id,
                            value = chitieu_11,
                            target_id = 11,
                            date = date,
                            created_at = DateTime.Now
                        };
                        list_Result.Add(result_11);
                        _context.Add(result_11);
                    }
                    else
                    {
                        find_result.value = chitieu_11;
                        _context.Update(find_result);
                        //_context.SaveChanges();
                    }
                    //_context.SaveChanges();
                }
                if (chitieu_12 != null)
                {
                    var find_result = _context.ResultModel.Where(d => d.target_id == 12 && d.date == date && d.point_id == findPoint.id).FirstOrDefault();
                    if (find_result == null)
                    {
                        var result_12 = new ResultModel()
                        {
                            point_id = findPoint.id,
                            value = chitieu_12,
                            target_id = 12,
                            date = date,
                            created_at = DateTime.Now
                        };
                        list_Result.Add(result_12);
                        _context.Add(result_12);
                    }
                    else
                    {
                        find_result.value = chitieu_12;
                        _context.Update(find_result);
                        //_context.SaveChanges();
                    }
                    //_context.SaveChanges();
                }

                if (chitieu_13 != null)
                {
                    var find_result = _context.ResultModel.Where(d => d.target_id == 13 && d.date == date && d.point_id == findPoint.id).FirstOrDefault();
                    if (find_result == null)
                    {
                        var result_13 = new ResultModel()
                        {
                            point_id = findPoint.id,
                            value = chitieu_13,
                            target_id = 13,
                            date = date,
                            created_at = DateTime.Now
                        };
                        list_Result.Add(result_13);
                        _context.Add(result_13);
                    }
                    else
                    {
                        find_result.value = chitieu_13;
                        _context.Update(find_result);
                        //_context.SaveChanges();
                    }
                    //_context.SaveChanges();
                }
                //EquipmentModel EquipmentModel = new EquipmentModel { code = code, name = name_vn, name_en = name_en, created_at = DateTime.Now };
                //_context.Add(EquipmentModel);
                //_context.SaveChanges();
            }
            _context.SaveChanges();

            return Ok(list_Result);
        }
        public async Task<IActionResult> importVitriVisinhKho()
        {
            return Ok();
            // Khởi tạo workbook để đọc
            Spire.Xls.Workbook book = new Spire.Xls.Workbook();
            book.LoadFromFile("./wwwroot/data/trend/visinh/Raw data for of microbial results of  equiment quality monitoring_Kho 2024.xlsx", ExcelVersion.Version2013);

            Spire.Xls.Worksheet sheet = book.Worksheets[0];
            var lastrow = sheet.LastDataRow;
            // nếu vẫn chưa gặp end thì vẫn lấy data
            Console.WriteLine(lastrow);
            var list_Result = new List<ResultModel>();
            var location_id = 1;
            var location = "";
            var stt = 0;
            for (int rowIndex = 7; rowIndex < lastrow; rowIndex++)
            {
                // lấy row hiện tại
                var nowRow = sheet.Rows[rowIndex];
                if (nowRow == null)
                    continue;
                // vì ta dùng 3 cột A, B, C => data của ta sẽ như sau
                //int numcount = nowRow.Cells.Count;
                //for(int y = 0;y<numcount - 1 ;y++)
                var code_vitri = nowRow.Cells[4] != null ? nowRow.Cells[4].Value : null;
                // Xuất ra thông tin lên màn hình
                Console.WriteLine("MS: {0} ", code_vitri);
                Console.WriteLine("rowIndex: {0} ", rowIndex);

                if (code_vitri == null)
                    continue;

                //var tansuat = 3;
                //DateTime? date = nowRow.Cells[2] != null ? nowRow.Cells[2].DateTimeValue : null;
                var tansuat_id = 3;
                var object_id = 2;
                var target = nowRow.Cells[1] != null ? nowRow.Cells[1].Value : null;
                var target_id = 7;
                switch (target)
                {
                    case "Active":
                        target_id = 7;
                        break;
                    case "Passive":
                        target_id = 8;
                        break;
                    case "Rodac":
                        target_id = 15;
                        break;
                }
                var name = nowRow.Cells[3] != null ? nowRow.Cells[3].Value : null;

                var list_name = name.Split("\r\n", StringSplitOptions.None);
                var name_vi = list_name[0];
                var name_en = list_name.Length > 1 ? list_name[1] : null;

                if (nowRow.Cells[0] != null && nowRow.Cells[0].Value != "")
                {
                    location = nowRow.Cells[0].Value;
                    var list_location = location.Split("\r\n", StringSplitOptions.None);
                    var location_vi = list_location[0];
                    var location_en = list_location.Length > 1 ? list_location[1] : null;

                    var findLocation = new LocationModel()
                    {
                        name = location_vi,
                        name_en = location_en,
                        parent = 11,
                        stt = stt++,
                        count_child = 0,
                        created_at = DateTime.Now,
                    };
                    _context.Add(findLocation);
                    _context.SaveChanges();
                    location_id = findLocation.id;
                }
                //var chitieu_9 = nowRow.Cells[3] != null && nowRow.Cells[3].Value != "NA" && nowRow.Cells[3].Value != "" ? nowRow.Cells[3].Value : null;
                //var chitieu_10 = nowRow.Cells[4] != null && nowRow.Cells[4].Value != "NA" && nowRow.Cells[4].Value != "" ? nowRow.Cells[4].Value : null;
                //decimal? chitieu_11 = nowRow.Cells[5] != null && nowRow.Cells[5].Value != "NA" && nowRow.Cells[5].Value != "" ? (decimal)nowRow.Cells[5].NumberValue : null;
                //decimal? chitieu_12 = nowRow.Cells[6] != null && nowRow.Cells[6].Value != "NA" && nowRow.Cells[6].Value != "" ? (decimal)nowRow.Cells[6].NumberValue : null;
                //decimal? chitieu_13 = nowRow.Cells[7] != null && nowRow.Cells[7].Value != "NA" && nowRow.Cells[7].Value != "" ? (decimal)nowRow.Cells[7].NumberValue : null;
                //var chitieu_14 = nowRow.Cells[8] != null && nowRow.Cells[8].Value != "NA" && nowRow.Cells[8].Value != "" ? nowRow.Cells[8].Value : null;

                var findPoint = _context.PointModel.Where(d => d.code == code_vitri).FirstOrDefault();
                if (findPoint == null)
                {
                    findPoint = new PointModel()
                    {
                        color = ColorTranslator.ToHtml(GenerateRandomColor(code_vitri)),
                        code = code_vitri,
                        name = name_vi,
                        name_en = name_en,
                        frequency_id = tansuat_id,
                        object_id = object_id,
                        location_id = location_id,
                        target_id = target_id,
                        created_at = DateTime.Now,
                    };
                    _context.Add(findPoint);
                    _context.SaveChanges();
                }



                //EquipmentModel EquipmentModel = new EquipmentModel { code = code, name = name_vn, name_en = name_en, created_at = DateTime.Now };
                //_context.Add(EquipmentModel);
                //_context.SaveChanges();
            }
            //_context.AddRange(list_Result);
            //_context.SaveChanges();

            return Ok(list_Result);
        }
        public async Task<IActionResult> importVitriVisinhRD()
        {
            return Ok();
            // Khởi tạo workbook để đọc
            Spire.Xls.Workbook book = new Spire.Xls.Workbook();
            book.LoadFromFile("./wwwroot/data/trend/visinh/Raw data for of microbial results of room quality monitoring RD _2024.xlsx", ExcelVersion.Version2013);

            Spire.Xls.Worksheet sheet = book.Worksheets[0];
            var lastrow = sheet.LastDataRow;
            // nếu vẫn chưa gặp end thì vẫn lấy data
            Console.WriteLine(lastrow);
            var list_Result = new List<ResultModel>();
            var location_id = 1;
            var location = "";
            var stt = 0;
            for (int rowIndex = 7; rowIndex < lastrow; rowIndex++)
            {
                // lấy row hiện tại
                var nowRow = sheet.Rows[rowIndex];
                if (nowRow == null)
                    continue;
                // vì ta dùng 3 cột A, B, C => data của ta sẽ như sau
                //int numcount = nowRow.Cells.Count;
                //for(int y = 0;y<numcount - 1 ;y++)
                var code_vitri = nowRow.Cells[4] != null ? nowRow.Cells[4].Value : null;
                // Xuất ra thông tin lên màn hình
                Console.WriteLine("MS: {0} ", code_vitri);
                Console.WriteLine("rowIndex: {0} ", rowIndex);

                if (code_vitri == null)
                    continue;

                //var tansuat = 3;
                //DateTime? date = nowRow.Cells[2] != null ? nowRow.Cells[2].DateTimeValue : null;
                var tansuat_id = 3;
                var object_id = 2;
                var target = nowRow.Cells[1] != null ? nowRow.Cells[1].Value : null;
                var target_id = 7;
                switch (target)
                {
                    case "Active":
                        target_id = 7;
                        break;
                    case "Passive":
                        target_id = 8;
                        break;
                    case "Rodac":
                        target_id = 15;
                        break;
                }
                var name = nowRow.Cells[3] != null ? nowRow.Cells[3].Value : null;

                var list_name = name.Split("\r\n", StringSplitOptions.None);
                var name_vi = list_name[0];
                var name_en = list_name.Length > 1 ? list_name[1] : null;

                if (nowRow.Cells[0] != null && nowRow.Cells[0].Value != "")
                {
                    location = nowRow.Cells[0].Value;
                    var list_location = location.Split("\r\n", StringSplitOptions.None);
                    var location_vi = list_location[0];
                    var location_en = list_location.Length > 1 ? list_location[1] : null;

                    var findLocation = new LocationModel()
                    {
                        name = location_vi,
                        name_en = location_en,
                        parent = 12,
                        stt = stt++,
                        count_child = 0,
                        created_at = DateTime.Now,
                    };
                    _context.Add(findLocation);
                    _context.SaveChanges();
                    location_id = findLocation.id;
                }
                //var chitieu_9 = nowRow.Cells[3] != null && nowRow.Cells[3].Value != "NA" && nowRow.Cells[3].Value != "" ? nowRow.Cells[3].Value : null;
                //var chitieu_10 = nowRow.Cells[4] != null && nowRow.Cells[4].Value != "NA" && nowRow.Cells[4].Value != "" ? nowRow.Cells[4].Value : null;
                //decimal? chitieu_11 = nowRow.Cells[5] != null && nowRow.Cells[5].Value != "NA" && nowRow.Cells[5].Value != "" ? (decimal)nowRow.Cells[5].NumberValue : null;
                //decimal? chitieu_12 = nowRow.Cells[6] != null && nowRow.Cells[6].Value != "NA" && nowRow.Cells[6].Value != "" ? (decimal)nowRow.Cells[6].NumberValue : null;
                //decimal? chitieu_13 = nowRow.Cells[7] != null && nowRow.Cells[7].Value != "NA" && nowRow.Cells[7].Value != "" ? (decimal)nowRow.Cells[7].NumberValue : null;
                //var chitieu_14 = nowRow.Cells[8] != null && nowRow.Cells[8].Value != "NA" && nowRow.Cells[8].Value != "" ? nowRow.Cells[8].Value : null;

                var findPoint = _context.PointModel.Where(d => d.code == code_vitri).FirstOrDefault();
                if (findPoint == null)
                {
                    findPoint = new PointModel()
                    {
                        color = ColorTranslator.ToHtml(GenerateRandomColor(code_vitri)),
                        code = code_vitri,
                        name = name_vi,
                        name_en = name_en,
                        frequency_id = tansuat_id,
                        object_id = object_id,
                        location_id = location_id,
                        target_id = target_id,
                        created_at = DateTime.Now,
                    };
                    _context.Add(findPoint);
                    _context.SaveChanges();
                }



                //EquipmentModel EquipmentModel = new EquipmentModel { code = code, name = name_vn, name_en = name_en, created_at = DateTime.Now };
                //_context.Add(EquipmentModel);
                //_context.SaveChanges();
            }
            //_context.AddRange(list_Result);
            //_context.SaveChanges();

            return Ok(list_Result);
        }
        public async Task<IActionResult> importVitriVisinhTPCN()
        {
            return Ok();
            // Khởi tạo workbook để đọc
            Spire.Xls.Workbook book = new Spire.Xls.Workbook();
            book.LoadFromFile("./wwwroot/data/trend/visinh/Raw data for of microbial results of room quality monitoring_TPCN- GRADE D-2024.xlsx", ExcelVersion.Version2013);

            Spire.Xls.Worksheet sheet = book.Worksheets[0];
            var lastrow = sheet.LastDataRow;
            // nếu vẫn chưa gặp end thì vẫn lấy data
            Console.WriteLine(lastrow);
            var list_Result = new List<ResultModel>();
            var location_id = 1;
            var location = "";
            var stt = 0;
            for (int rowIndex = 7; rowIndex < lastrow; rowIndex++)
            {
                // lấy row hiện tại
                var nowRow = sheet.Rows[rowIndex];
                if (nowRow == null)
                    continue;
                // vì ta dùng 3 cột A, B, C => data của ta sẽ như sau
                //int numcount = nowRow.Cells.Count;
                //for(int y = 0;y<numcount - 1 ;y++)
                var code_vitri = nowRow.Cells[4] != null ? nowRow.Cells[4].Value : null;
                // Xuất ra thông tin lên màn hình
                Console.WriteLine("MS: {0} ", code_vitri);
                Console.WriteLine("rowIndex: {0} ", rowIndex);

                if (code_vitri == null)
                    continue;

                var tansuat = nowRow.Cells[2] != null ? nowRow.Cells[2].Value : null;
                //DateTime? date = nowRow.Cells[2] != null ? nowRow.Cells[2].DateTimeValue : null;
                var tansuat_id = 3;
                switch (tansuat)
                {
                    case "6 tháng/ lần":
                        tansuat_id = 5;
                        break;
                }
                var object_id = 2;
                var target = nowRow.Cells[1] != null ? nowRow.Cells[1].Value : null;
                var target_id = 7;
                switch (target)
                {
                    case "Active":
                        target_id = 7;
                        break;
                    case "Passive":
                        target_id = 8;
                        break;
                    case "Rodac":
                        target_id = 15;
                        break;
                }
                var name = nowRow.Cells[3] != null ? nowRow.Cells[3].Value : null;

                var list_name = name.Split("\r\n", StringSplitOptions.None);
                var name_vi = list_name[0];
                var name_en = list_name.Length > 1 ? list_name[1] : null;

                if (nowRow.Cells[0] != null && nowRow.Cells[0].Value != "")
                {
                    location = nowRow.Cells[0].Value;
                    var list_location = location.Split("\r\n", StringSplitOptions.None);
                    var location_vi = list_location[0];
                    var location_en = list_location.Length > 1 ? list_location[1] : null;

                    var findLocation = new LocationModel()
                    {
                        color = ColorTranslator.ToHtml(GenerateRandomColor(code_vitri)),
                        name = location_vi,
                        name_en = location_en,
                        parent = 4,
                        stt = stt++,
                        count_child = 0,
                        created_at = DateTime.Now,
                    };
                    _context.Add(findLocation);
                    _context.SaveChanges();
                    location_id = findLocation.id;
                }
                //var chitieu_9 = nowRow.Cells[3] != null && nowRow.Cells[3].Value != "NA" && nowRow.Cells[3].Value != "" ? nowRow.Cells[3].Value : null;
                //var chitieu_10 = nowRow.Cells[4] != null && nowRow.Cells[4].Value != "NA" && nowRow.Cells[4].Value != "" ? nowRow.Cells[4].Value : null;
                //decimal? chitieu_11 = nowRow.Cells[5] != null && nowRow.Cells[5].Value != "NA" && nowRow.Cells[5].Value != "" ? (decimal)nowRow.Cells[5].NumberValue : null;
                //decimal? chitieu_12 = nowRow.Cells[6] != null && nowRow.Cells[6].Value != "NA" && nowRow.Cells[6].Value != "" ? (decimal)nowRow.Cells[6].NumberValue : null;
                //decimal? chitieu_13 = nowRow.Cells[7] != null && nowRow.Cells[7].Value != "NA" && nowRow.Cells[7].Value != "" ? (decimal)nowRow.Cells[7].NumberValue : null;
                //var chitieu_14 = nowRow.Cells[8] != null && nowRow.Cells[8].Value != "NA" && nowRow.Cells[8].Value != "" ? nowRow.Cells[8].Value : null;

                var findPoint = _context.PointModel.Where(d => d.code == code_vitri).FirstOrDefault();
                if (findPoint == null)
                {
                    findPoint = new PointModel()
                    {
                        color = ColorTranslator.ToHtml(GenerateRandomColor(code_vitri)),
                        code = code_vitri,
                        name = name_vi,
                        name_en = name_en,
                        frequency_id = tansuat_id,
                        object_id = object_id,
                        location_id = location_id,
                        target_id = target_id,
                        created_at = DateTime.Now,
                    };
                    _context.Add(findPoint);
                    _context.SaveChanges();
                }



                //EquipmentModel EquipmentModel = new EquipmentModel { code = code, name = name_vn, name_en = name_en, created_at = DateTime.Now };
                //_context.Add(EquipmentModel);
                //_context.SaveChanges();
            }
            //_context.AddRange(list_Result);
            //_context.SaveChanges();

            return Ok(list_Result);
        }

        public async Task<IActionResult> importVitriVisinhQCD()
        {
            return Ok();
            // Khởi tạo workbook để đọc
            Spire.Xls.Workbook book = new Spire.Xls.Workbook();
            book.LoadFromFile("./wwwroot/data/trend/visinh/Raw data for of microbial results of room quality monitoring_QC- GRADE D.xlsx", ExcelVersion.Version2013);

            Spire.Xls.Worksheet sheet = book.Worksheets[1];
            var lastrow = sheet.LastDataRow;
            // nếu vẫn chưa gặp end thì vẫn lấy data
            Console.WriteLine(lastrow);
            var list_Result = new List<ResultModel>();
            var location_id = 1;
            var location = "";
            var stt = 0;
            for (int rowIndex = 8; rowIndex < lastrow; rowIndex++)
            {
                // lấy row hiện tại
                var nowRow = sheet.Rows[rowIndex];
                if (nowRow == null)
                    continue;
                // vì ta dùng 3 cột A, B, C => data của ta sẽ như sau
                //int numcount = nowRow.Cells.Count;
                //for(int y = 0;y<numcount - 1 ;y++)
                var code_vitri = nowRow.Cells[4] != null ? nowRow.Cells[4].Value : null;
                // Xuất ra thông tin lên màn hình
                Console.WriteLine("MS: {0} ", code_vitri);
                Console.WriteLine("rowIndex: {0} ", rowIndex);

                if (code_vitri == null)
                    continue;

                var tansuat = nowRow.Cells[2] != null ? nowRow.Cells[2].Value : null;
                //DateTime? date = nowRow.Cells[2] != null ? nowRow.Cells[2].DateTimeValue : null;
                var tansuat_id = 3;
                switch (tansuat)
                {
                    case "6 tháng/ lần":
                        tansuat_id = 5;
                        break;
                    case "2 tuần/ lần":
                        tansuat_id = 2;
                        break;
                }
                var object_id = 2;
                var target = nowRow.Cells[1] != null ? nowRow.Cells[1].Value : null;
                var target_id = 7;
                switch (target)
                {
                    case "Active":
                        target_id = 7;
                        break;
                    case "Passive":
                        target_id = 8;
                        break;
                    case "Rodac":
                        target_id = 15;
                        break;
                }
                var name = nowRow.Cells[3] != null ? nowRow.Cells[3].Value : null;

                var list_name = name.Split("\r\n", StringSplitOptions.None);
                var name_vi = list_name[0];
                var name_en = list_name.Length > 1 ? list_name[1] : null;

                if (nowRow.Cells[0] != null && nowRow.Cells[0].Value != "")
                {
                    location = nowRow.Cells[0].Value;
                    var list_location = location.Split("\r\n", StringSplitOptions.None);
                    var location_vi = list_location[0];
                    var location_en = list_location.Length > 1 ? list_location[1] : null;

                    var findLocation = new LocationModel()
                    {
                        color = ColorTranslator.ToHtml(GenerateRandomColor(code_vitri)),
                        name = location_vi,
                        name_en = location_en,
                        parent = 9,
                        stt = stt++,
                        count_child = 0,
                        created_at = DateTime.Now,
                    };
                    _context.Add(findLocation);
                    _context.SaveChanges();
                    location_id = findLocation.id;
                }
                //var chitieu_9 = nowRow.Cells[3] != null && nowRow.Cells[3].Value != "NA" && nowRow.Cells[3].Value != "" ? nowRow.Cells[3].Value : null;
                //var chitieu_10 = nowRow.Cells[4] != null && nowRow.Cells[4].Value != "NA" && nowRow.Cells[4].Value != "" ? nowRow.Cells[4].Value : null;
                //decimal? chitieu_11 = nowRow.Cells[5] != null && nowRow.Cells[5].Value != "NA" && nowRow.Cells[5].Value != "" ? (decimal)nowRow.Cells[5].NumberValue : null;
                //decimal? chitieu_12 = nowRow.Cells[6] != null && nowRow.Cells[6].Value != "NA" && nowRow.Cells[6].Value != "" ? (decimal)nowRow.Cells[6].NumberValue : null;
                //decimal? chitieu_13 = nowRow.Cells[7] != null && nowRow.Cells[7].Value != "NA" && nowRow.Cells[7].Value != "" ? (decimal)nowRow.Cells[7].NumberValue : null;
                //var chitieu_14 = nowRow.Cells[8] != null && nowRow.Cells[8].Value != "NA" && nowRow.Cells[8].Value != "" ? nowRow.Cells[8].Value : null;

                var findPoint = _context.PointModel.Where(d => d.code == code_vitri).FirstOrDefault();
                if (findPoint == null)
                {
                    findPoint = new PointModel()
                    {
                        color = ColorTranslator.ToHtml(GenerateRandomColor(code_vitri)),
                        code = code_vitri,
                        name = name_vi,
                        name_en = name_en,
                        frequency_id = tansuat_id,
                        object_id = object_id,
                        location_id = location_id,
                        target_id = target_id,
                        created_at = DateTime.Now,
                    };
                    _context.Add(findPoint);
                    _context.SaveChanges();
                }



                //EquipmentModel EquipmentModel = new EquipmentModel { code = code, name = name_vn, name_en = name_en, created_at = DateTime.Now };
                //_context.Add(EquipmentModel);
                //_context.SaveChanges();
            }
            //_context.AddRange(list_Result);
            //_context.SaveChanges();

            return Ok(list_Result);
        }

        public async Task<IActionResult> importVitriVisinhQCC()
        {
            return Ok();
            // Khởi tạo workbook để đọc
            Spire.Xls.Workbook book = new Spire.Xls.Workbook();
            book.LoadFromFile("./wwwroot/data/trend/visinh/Raw data for of microbial results of  equiment quality monitoring_QC- GRADE C.xlsx", ExcelVersion.Version2013);

            Spire.Xls.Worksheet sheet = book.Worksheets[1];
            var lastrow = sheet.LastDataRow;
            // nếu vẫn chưa gặp end thì vẫn lấy data
            Console.WriteLine(lastrow);
            var list_Result = new List<ResultModel>();
            var location_id = 1;
            var location = "";
            var stt = 0;
            for (int rowIndex = 8; rowIndex < lastrow; rowIndex++)
            {
                // lấy row hiện tại
                var nowRow = sheet.Rows[rowIndex];
                if (nowRow == null)
                    continue;
                // vì ta dùng 3 cột A, B, C => data của ta sẽ như sau
                //int numcount = nowRow.Cells.Count;
                //for(int y = 0;y<numcount - 1 ;y++)
                var code_vitri = nowRow.Cells[4] != null ? nowRow.Cells[4].Value : null;
                // Xuất ra thông tin lên màn hình
                Console.WriteLine("MS: {0} ", code_vitri);
                Console.WriteLine("rowIndex: {0} ", rowIndex);

                if (code_vitri == null)
                    continue;

                var tansuat = nowRow.Cells[2] != null ? nowRow.Cells[2].Value : null;
                //DateTime? date = nowRow.Cells[2] != null ? nowRow.Cells[2].DateTimeValue : null;
                var tansuat_id = 3;
                switch (tansuat)
                {
                    case "6 tháng/ lần":
                        tansuat_id = 5;
                        break;
                    case "2 tuần/ lần":
                        tansuat_id = 2;
                        break;
                }
                var object_id = 2;
                var target = nowRow.Cells[1] != null ? nowRow.Cells[1].Value : null;
                var target_id = 7;
                switch (target)
                {
                    case "Active":
                        target_id = 7;
                        break;
                    case "Passive":
                        target_id = 8;
                        break;
                    case "Rodac":
                        target_id = 15;
                        break;
                }
                var name = nowRow.Cells[3] != null ? nowRow.Cells[3].Value : null;

                var list_name = name.Split("\r\n", StringSplitOptions.None);
                var name_vi = list_name[0];
                var name_en = list_name.Length > 1 ? list_name[1] : null;

                if (nowRow.Cells[0] != null && nowRow.Cells[0].Value != "")
                {
                    location = nowRow.Cells[0].Value;
                    var list_location = location.Split("\r\n", StringSplitOptions.None);
                    var location_vi = list_location[0];
                    var location_en = list_location.Length > 1 ? list_location[1] : null;

                    var findLocation = new LocationModel()
                    {
                        name = location_vi,
                        name_en = location_en,
                        parent = 10,
                        stt = stt++,
                        count_child = 0,
                        created_at = DateTime.Now,
                    };
                    _context.Add(findLocation);
                    _context.SaveChanges();
                    location_id = findLocation.id;
                }
                //var chitieu_9 = nowRow.Cells[3] != null && nowRow.Cells[3].Value != "NA" && nowRow.Cells[3].Value != "" ? nowRow.Cells[3].Value : null;
                //var chitieu_10 = nowRow.Cells[4] != null && nowRow.Cells[4].Value != "NA" && nowRow.Cells[4].Value != "" ? nowRow.Cells[4].Value : null;
                //decimal? chitieu_11 = nowRow.Cells[5] != null && nowRow.Cells[5].Value != "NA" && nowRow.Cells[5].Value != "" ? (decimal)nowRow.Cells[5].NumberValue : null;
                //decimal? chitieu_12 = nowRow.Cells[6] != null && nowRow.Cells[6].Value != "NA" && nowRow.Cells[6].Value != "" ? (decimal)nowRow.Cells[6].NumberValue : null;
                //decimal? chitieu_13 = nowRow.Cells[7] != null && nowRow.Cells[7].Value != "NA" && nowRow.Cells[7].Value != "" ? (decimal)nowRow.Cells[7].NumberValue : null;
                //var chitieu_14 = nowRow.Cells[8] != null && nowRow.Cells[8].Value != "NA" && nowRow.Cells[8].Value != "" ? nowRow.Cells[8].Value : null;

                var findPoint = _context.PointModel.Where(d => d.code == code_vitri).FirstOrDefault();
                if (findPoint == null)
                {
                    findPoint = new PointModel()
                    {
                        color = ColorTranslator.ToHtml(GenerateRandomColor(code_vitri)),
                        code = code_vitri,
                        name = name_vi,
                        name_en = name_en,
                        frequency_id = tansuat_id,
                        object_id = object_id,
                        location_id = location_id,
                        target_id = target_id,
                        created_at = DateTime.Now,
                    };
                    _context.Add(findPoint);
                    _context.SaveChanges();
                }



                //EquipmentModel EquipmentModel = new EquipmentModel { code = code, name = name_vn, name_en = name_en, created_at = DateTime.Now };
                //_context.Add(EquipmentModel);
                //_context.SaveChanges();
            }
            //_context.AddRange(list_Result);
            //_context.SaveChanges();

            return Ok(list_Result);
        }

        public async Task<IActionResult> importVitriDungngoai()
        {
            return Ok();
            // Khởi tạo workbook để đọc
            Spire.Xls.Workbook book = new Spire.Xls.Workbook();
            book.LoadFromFile("./wwwroot/data/trend/visinh/EM XUONG DUNG NGOAI (Thường quy).xlsx", ExcelVersion.Version2013);
            var worksheets = book.Worksheets.Count();
            var list_Result = new List<ResultModel>();
            for (var worksheetsIndex = 0; worksheetsIndex < worksheets; worksheetsIndex++)
            {
                Spire.Xls.Worksheet sheet = book.Worksheets[worksheetsIndex];
                var lastrow = sheet.LastDataRow;
                var lastcol = sheet.LastDataColumn;
                // nếu vẫn chưa gặp end thì vẫn lấy data
                Console.WriteLine(lastrow);
                var location_id = 1;
                var location = "";
                var stt = 0;
                for (int rowIndex = 1; rowIndex < lastrow; rowIndex++)
                {
                    // lấy row hiện tại
                    Console.WriteLine("rowIndex: {0} ", rowIndex);
                    var nowRow = sheet.Rows[rowIndex];
                    if (nowRow == null)
                        continue;

                    string? code = nowRow.Cells[4] != null && nowRow.Cells[4].Value != "" ? nowRow.Cells[4].Value : null;
                    if (code == null)
                        continue;
                    Console.WriteLine("date: {0} ", code);
                    int? target_id = 7;
                    if (code.Contains('A'))
                    {
                        target_id = 7;
                    }
                    else if (code.Contains('P'))
                    {
                        target_id = 8;
                    }
                    else if (code.Contains('R'))
                    {
                        target_id = 15;

                    }
                    if (nowRow.Cells[1] != null && nowRow.Cells[1].Value != "")
                    {
                        location = nowRow.Cells[1].Value;
                        var list_location = location.Split("\r\n", StringSplitOptions.None);
                        var location_vi = list_location[0];
                        var location_en = list_location.Length > 1 ? list_location[1] : null;
                        var code_location = nowRow.Cells[2].Value;
                        var findLocation = _context.LocationModel.Where(d => d.name == location_vi && d.parent == 1444).FirstOrDefault();

                        if (findLocation == null)
                        {
                            findLocation = new LocationModel()
                            {
                                name = location_vi,
                                name_en = location_en,
                                code = code_location,
                                parent = 1444,
                                stt = stt++,
                                count_child = 0,
                                created_at = DateTime.Now,
                            };
                            _context.Add(findLocation);
                            _context.SaveChanges();
                        }
                        else
                        {
                            findLocation.code = code_location;
                            findLocation.name_en = location_en;
                            _context.Update(findLocation);
                            _context.SaveChanges();
                        }
                        location_id = findLocation.id;
                    }

                    string? name = nowRow.Cells[3] != null && nowRow.Cells[3].Value != "" ? nowRow.Cells[3].Value : null;

                    var findPoint = _context.PointModel.Where(d => d.code == code).FirstOrDefault();
                    if (findPoint == null)
                    {
                        findPoint = new PointModel()
                        {
                            color = ColorTranslator.ToHtml(GenerateRandomColor(code)),
                            code = code,
                            name = name,
                            frequency_id = 3,
                            object_id = 2,
                            location_id = location_id,
                            target_id = target_id,
                            created_at = DateTime.Now,
                        };
                        _context.Add(findPoint);
                        _context.SaveChanges();
                    }

                }
            }


            return Ok(list_Result);
        }
        public async Task<IActionResult> importVitriNON()
        {
            return Ok();
            // Khởi tạo workbook để đọc
            Spire.Xls.Workbook book = new Spire.Xls.Workbook();
            book.LoadFromFile("./wwwroot/data/trend/visinh/EM XUONG NON.xlsx", ExcelVersion.Version2013);
            var worksheets = book.Worksheets.Count();
            var list_Result = new List<ResultModel>();
            for (var worksheetsIndex = 0; worksheetsIndex < worksheets; worksheetsIndex++)
            {
                Spire.Xls.Worksheet sheet = book.Worksheets[worksheetsIndex];
                var lastrow = sheet.LastDataRow;
                var lastcol = sheet.LastDataColumn;
                // nếu vẫn chưa gặp end thì vẫn lấy data
                Console.WriteLine(lastrow);
                var location_id = 1;
                var location = "";
                var stt = 0;
                for (int rowIndex = 2; rowIndex < lastrow; rowIndex++)
                {
                    // lấy row hiện tại
                    Console.WriteLine("rowIndex: {0} ", rowIndex);
                    var nowRow = sheet.Rows[rowIndex];
                    if (nowRow == null)
                        continue;

                    string? code = nowRow.Cells[2] != null && nowRow.Cells[2].Value != "" ? nowRow.Cells[2].Value : null;
                    if (code == null)
                        continue;
                    Console.WriteLine("date: {0} ", code);
                    int? target_id = 7;
                    if (code.Contains('A'))
                    {
                        target_id = 7;
                    }
                    else if (code.Contains('P'))
                    {
                        target_id = 8;
                    }
                    else if (code.Contains('R'))
                    {
                        target_id = 15;

                    }
                    if (nowRow.Cells[0] != null && nowRow.Cells[0].Value != "")
                    {
                        location = nowRow.Cells[0].Value;
                        var list_location = location.Split("\r\n", StringSplitOptions.None);
                        var location_vi = list_location[0];
                        var location_en = list_location.Length > 1 ? list_location[1] : null;
                        var findLocation = _context.LocationModel.Where(d => d.name == location_vi && d.parent == 1443).FirstOrDefault();

                        if (findLocation == null)
                        {
                            findLocation = new LocationModel()
                            {
                                name = location_vi,
                                name_en = location_en,
                                parent = 1443,
                                stt = stt++,
                                count_child = 0,
                                created_at = DateTime.Now,
                            };
                            _context.Add(findLocation);
                            _context.SaveChanges();
                        }
                        else
                        {
                            findLocation.name_en = location_en;
                            _context.Update(findLocation);
                            _context.SaveChanges();
                        }
                        location_id = findLocation.id;
                    }

                    string? name = nowRow.Cells[1] != null && nowRow.Cells[1].Value != "" ? nowRow.Cells[1].Value : null;

                    var findPoint = _context.PointModel.Where(d => d.code == code).FirstOrDefault();
                    if (findPoint == null)
                    {
                        findPoint = new PointModel()
                        {
                            color = ColorTranslator.ToHtml(GenerateRandomColor(code)),
                            code = code,
                            name = name,
                            frequency_id = 3,
                            object_id = 2,
                            location_id = location_id,
                            target_id = target_id,
                            created_at = DateTime.Now,
                        };
                        _context.Add(findPoint);
                        _context.SaveChanges();
                    }

                }
            }


            return Ok(list_Result);
        }

        public async Task<IActionResult> importCA2()
        {
            return Ok();
            // Khởi tạo workbook để đọc
            Spire.Xls.Workbook book = new Spire.Xls.Workbook();
            book.LoadFromFile("./wwwroot/data/trend/khinen/CA_23_05_24.xlsx", ExcelVersion.Version2013);

            Spire.Xls.Worksheet sheet = book.Worksheets[0];
            var lastrow = sheet.LastDataRow;
            // nếu vẫn chưa gặp end thì vẫn lấy data
            Console.WriteLine(lastrow);
            var list_Result = new List<ResultModel>();
            for (int rowIndex = 2; rowIndex < lastrow; rowIndex++)
            {
                // lấy row hiện tại
                var nowRow = sheet.Rows[rowIndex];
                if (nowRow == null)
                    continue;
                // vì ta dùng 3 cột A, B, C => data của ta sẽ như sau
                //int numcount = nowRow.Cells.Count;
                //for(int y = 0;y<numcount - 1 ;y++)
                var code_vitri = nowRow.Cells[4] != null ? nowRow.Cells[4].Value : null;
                // Xuất ra thông tin lên màn hình
                Console.WriteLine("MS: {0} ", code_vitri);
                Console.WriteLine("rowIndex: {0} ", rowIndex);

                if (code_vitri == null)
                    continue;

                //var tansuat = nowRow.Cells[2] != null ? nowRow.Cells[2].Value : null;
                //DateTime? date = nowRow.Cells[3] != null ? nowRow.Cells[3].DateTimeValue : null;
                var tansuat_id = 4;
                //switch (tansuat)
                //{
                //    case "Daily":
                //        tansuat_id = 1;
                //        break;
                //    case "Twice per month":
                //        tansuat_id = 2;
                //        break;

                //}
                var phong = nowRow.Cells[1] != null ? nowRow.Cells[1].Value : null;
                var list_name = phong.Split("\r\n", StringSplitOptions.None);
                var khuvuc = nowRow.Cells[2] != null ? nowRow.Cells[2].Value : null;
                var parent = 0;
                switch (khuvuc)
                {
                    case "Xưởng dùng ngoài":
                        parent = 1444;
                        break;
                    case "Xưởng thuốc uống Non-Betalactam":
                        parent = 1443;
                        break;

                }
                var department = list_name[0];
                var department_en = list_name[1];
                var finddepartment = _context.LocationModel.Where(d => d.name == department && d.parent == parent).FirstOrDefault();
                if (finddepartment == null)
                {
                    finddepartment = new LocationModel()
                    {
                        name = department,
                        name_en = department_en,
                        parent = parent,
                        stt = 0,
                        count_child = 0,
                        created_at = DateTime.Now,
                    };
                    _context.Add(finddepartment);
                    _context.SaveChanges();
                }
                var findPoint = _context.PointModel.Where(d => d.code == code_vitri).FirstOrDefault();
                if (findPoint == null)
                {
                    findPoint = new PointModel()
                    {
                        color = ColorTranslator.ToHtml(GenerateRandomColor(code_vitri)),
                        code = code_vitri,
                        frequency_id = tansuat_id,
                        object_id = 4,
                        location_id = finddepartment.id,
                        created_at = DateTime.Now,
                    };
                    _context.Add(findPoint);
                    _context.SaveChanges();
                }

                //EquipmentModel EquipmentModel = new EquipmentModel { code = code, name = name_vn, name_en = name_en, created_at = DateTime.Now };
                //_context.Add(EquipmentModel);
                //_context.SaveChanges();
            }
            _context.AddRange(list_Result);
            _context.SaveChanges();

            return Ok(list_Result);
        }
        public async Task<IActionResult> importCA()
        {
            return Ok();
            // Khởi tạo workbook để đọc
            Spire.Xls.Workbook book = new Spire.Xls.Workbook();
            book.LoadFromFile("./wwwroot/data/trend/khinen/CA_26_04_24.xlsx", ExcelVersion.Version2013);

            Spire.Xls.Worksheet sheet = book.Worksheets[0];
            var lastrow = sheet.LastDataRow;
            // nếu vẫn chưa gặp end thì vẫn lấy data
            Console.WriteLine(lastrow);
            var list_Result = new List<ResultModel>();
            for (int rowIndex = 1; rowIndex < lastrow; rowIndex++)
            {
                // lấy row hiện tại
                var nowRow = sheet.Rows[rowIndex];
                if (nowRow == null)
                    continue;
                // vì ta dùng 3 cột A, B, C => data của ta sẽ như sau
                //int numcount = nowRow.Cells.Count;
                //for(int y = 0;y<numcount - 1 ;y++)
                var code_vitri = nowRow.Cells[1] != null ? nowRow.Cells[1].Value : null;
                // Xuất ra thông tin lên màn hình
                Console.WriteLine("MS: {0} ", code_vitri);
                Console.WriteLine("rowIndex: {0} ", rowIndex);

                if (code_vitri == null)
                    continue;

                //var tansuat = nowRow.Cells[2] != null ? nowRow.Cells[2].Value : null;
                DateTime? date = nowRow.Cells[3] != null ? nowRow.Cells[3].DateTimeValue : null;
                var tansuat_id = 4;
                //switch (tansuat)
                //{
                //    case "Daily":
                //        tansuat_id = 1;
                //        break;
                //    case "Twice per month":
                //        tansuat_id = 2;
                //        break;

                //}
                var phong = nowRow.Cells[0] != null ? nowRow.Cells[0].Value : null;
                var list_name = phong.Split("\r\n", StringSplitOptions.None);
                var khuvuc = list_name[2];
                var parent = 0;
                switch (khuvuc)
                {
                    case "(HSW)":
                        parent = 442;
                        break;
                    case "(R&D)":
                        parent = 440;
                        break;

                }
                var department = list_name[0];
                var department_en = list_name[1];
                var finddepartment = _context.LocationModel.Where(d => d.name == department && d.parent == parent).FirstOrDefault();
                if (finddepartment == null)
                {
                    finddepartment = new LocationModel()
                    {
                        name = department,
                        name_en = department_en,
                        parent = parent,
                        stt = 0,
                        count_child = 0,
                        created_at = DateTime.Now,
                    };
                    _context.Add(finddepartment);
                    _context.SaveChanges();
                }
                var findPoint = _context.PointModel.Where(d => d.code == code_vitri).FirstOrDefault();
                if (findPoint == null)
                {
                    findPoint = new PointModel()
                    {
                        color = ColorTranslator.ToHtml(GenerateRandomColor(code_vitri)),
                        code = code_vitri,
                        frequency_id = tansuat_id,
                        object_id = 4,
                        location_id = finddepartment.id,
                        created_at = DateTime.Now,
                    };
                    _context.Add(findPoint);
                    _context.SaveChanges();
                }

                //EquipmentModel EquipmentModel = new EquipmentModel { code = code, name = name_vn, name_en = name_en, created_at = DateTime.Now };
                //_context.Add(EquipmentModel);
                //_context.SaveChanges();
            }
            _context.AddRange(list_Result);
            _context.SaveChanges();

            return Ok(list_Result);
        }
        public async Task<IActionResult> dataCA()
        {
            return Ok();
            // Khởi tạo workbook để đọc
            Spire.Xls.Workbook book = new Spire.Xls.Workbook();
            book.LoadFromFile("./wwwroot/data/trend/khinen/Raw data for CA system quality monitoring_2024.xlsx", ExcelVersion.Version2013);

            Spire.Xls.Worksheet sheet = book.Worksheets[0];
            var lastrow = sheet.LastDataRow;
            // nếu vẫn chưa gặp end thì vẫn lấy data
            Console.WriteLine(lastrow);
            var list_Result = new List<ResultModel>();
            for (int rowIndex = 7; rowIndex < lastrow; rowIndex++)
            {
                // lấy row hiện tại
                var nowRow = sheet.Rows[rowIndex];
                if (nowRow == null)
                    continue;
                // vì ta dùng 3 cột A, B, C => data của ta sẽ như sau
                //int numcount = nowRow.Cells.Count;
                //for(int y = 0;y<numcount - 1 ;y++)
                var code_vitri = nowRow.Cells[1] != null ? nowRow.Cells[1].Value : null;
                // Xuất ra thông tin lên màn hình
                Console.WriteLine("MS: {0} ", code_vitri);
                Console.WriteLine("rowIndex: {0} ", rowIndex);

                if (code_vitri == null)
                    continue;

                //var tansuat = nowRow.Cells[2] != null ? nowRow.Cells[2].Value : null;
                DateTime? date = nowRow.Cells[3] != null ? nowRow.Cells[3].DateTimeValue : null;
                var tansuat_id = 4;
                //switch (tansuat)
                //{
                //    case "Daily":
                //        tansuat_id = 1;
                //        break;
                //    case "Twice per month":
                //        tansuat_id = 2;
                //        break;

                //}
                var findPoint = _context.PointModel.Where(d => d.code == code_vitri).AsNoTracking().FirstOrDefault();
                if (findPoint == null)
                {
                    continue;
                }
                var chitieu_16 = nowRow.Cells[4] != null && nowRow.Cells[4].Value != "NA" && nowRow.Cells[4].Value != "" ? nowRow.Cells[4].Value : null;
                var chitieu_17 = nowRow.Cells[5] != null && nowRow.Cells[5].Value != "NA" && nowRow.Cells[5].Value != "" ? nowRow.Cells[5].Value : null;
                decimal? chitieu_18 = nowRow.Cells[6] != null && nowRow.Cells[6].Value != "NA" && nowRow.Cells[6].Value != "" ? (decimal)nowRow.Cells[6].NumberValue : null;
                decimal? chitieu_19 = nowRow.Cells[7] != null && nowRow.Cells[7].Value != "NA" && nowRow.Cells[7].Value != "" ? (decimal)nowRow.Cells[7].NumberValue : null;
                decimal? chitieu_20 = nowRow.Cells[8] != null && nowRow.Cells[8].Value != "NA" && nowRow.Cells[8].Value != "" ? (decimal)nowRow.Cells[8].NumberValue : null;


                if (chitieu_16 != null)
                {
                    var chitieu_16_int = chitieu_16.ToUpper() == "ĐẠT" ? 1 : 0;
                    var find_result = _context.ResultModel.Where(d => d.target_id == 16 && d.date == date && d.point_id == findPoint.id).FirstOrDefault();
                    if (find_result == null)
                    {

                        var result_16 = new ResultModel()
                        {
                            point_id = findPoint.id,
                            value = chitieu_16_int,
                            object_id = 4,
                            target_id = 16,
                            date = date,
                            created_at = DateTime.Now
                        };
                        //list_Result.Add(result_16);
                        _context.Add(result_16);
                        _context.SaveChanges();
                    }
                    else
                    {
                        find_result.value = chitieu_16_int;
                        _context.Update(find_result);
                        _context.SaveChanges();
                    }
                }
                if (chitieu_17 != null)
                {
                    var chitieu_17_int = chitieu_17.ToUpper() == "ĐẠT" ? 1 : 0;
                    var find_result = _context.ResultModel.Where(d => d.target_id == 17 && d.date == date && d.point_id == findPoint.id).FirstOrDefault();
                    if (find_result == null)
                    {
                        var result_17 = new ResultModel()
                        {
                            point_id = findPoint.id,
                            value = chitieu_17_int,
                            object_id = 4,
                            target_id = 17,
                            date = date,
                            created_at = DateTime.Now
                        };
                        //list_Result.Add(result_17);

                        _context.Add(result_17);
                        _context.SaveChanges();
                    }
                    else
                    {
                        find_result.value = chitieu_17_int;
                        _context.Update(find_result);
                        _context.SaveChanges();
                    }
                    //_context.SaveChanges();
                }

                if (chitieu_18 != null)
                {
                    var find_result = _context.ResultModel.Where(d => d.target_id == 18 && d.date == date && d.point_id == findPoint.id).FirstOrDefault();
                    if (find_result == null)
                    {
                        var result_18 = new ResultModel()
                        {
                            point_id = findPoint.id,
                            value = chitieu_18,
                            object_id = 4,
                            target_id = 18,
                            date = date,
                            created_at = DateTime.Now
                        };
                        list_Result.Add(result_18);
                        _context.Add(result_18);
                        _context.SaveChanges();
                    }
                    else
                    {
                        find_result.value = chitieu_18;
                        _context.Update(find_result);
                        _context.SaveChanges();
                    }
                    //_context.SaveChanges();
                }
                if (chitieu_19 != null)
                {
                    var find_result = _context.ResultModel.Where(d => d.target_id == 19 && d.date == date && d.point_id == findPoint.id).FirstOrDefault();
                    if (find_result == null)
                    {
                        var result_19 = new ResultModel()
                        {
                            point_id = findPoint.id,
                            value = chitieu_19,
                            object_id = 4,
                            target_id = 19,
                            date = date,
                            created_at = DateTime.Now
                        };
                        //list_Result.Add(result_19);
                        _context.Add(result_19);
                        _context.SaveChanges();
                    }
                    else
                    {
                        find_result.value = chitieu_19;
                        _context.Update(find_result);
                        _context.SaveChanges();
                    }
                    //_context.SaveChanges();
                }

                if (chitieu_20 != null)
                {
                    var find_result = _context.ResultModel.Where(d => d.target_id == 20 && d.date == date && d.point_id == findPoint.id).FirstOrDefault();
                    if (find_result == null)
                    {
                        var result_20 = new ResultModel()
                        {
                            point_id = findPoint.id,
                            value = chitieu_20,
                            object_id = 4,
                            target_id = 20,
                            date = date,
                            created_at = DateTime.Now
                        };
                        //list_Result.Add(result_20);
                        _context.Add(result_20);
                        _context.SaveChanges();
                    }
                    else
                    {
                        find_result.value = chitieu_20;
                        _context.Update(find_result);
                        _context.SaveChanges();
                    }
                    //_context.SaveChanges();
                }
                //EquipmentModel EquipmentModel = new EquipmentModel { code = code, name = name_vn, name_en = name_en, created_at = DateTime.Now };
                //_context.Add(EquipmentModel);
                //_context.SaveChanges();
            }
            //_context.AddRange(list_Result);
            //_context.SaveChanges();

            return Json(list_Result, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }

        public async Task<IActionResult> dataVisinhKho()
        {
            //return Ok();
            // Khởi tạo workbook để đọc
            Spire.Xls.Workbook book = new Spire.Xls.Workbook();
            //book.LoadFromFile("./wwwroot/data/trend/visinh/Raw data for of microbial results of  equiment quality monitoring_Kho 2023.xlsx", ExcelVersion.Version2013);
            //book.LoadFromFile("./wwwroot/data/trend/visinh/Raw data for of microbial results of  equiment quality monitoring_Kho 2024.xlsx", ExcelVersion.Version2013);
            book.LoadFromFile("./wwwroot/data/trend/visinh/Raw/Raw data for of microbial results of  equiment quality monitoring_Kho 2024.xlsx", ExcelVersion.Version2013);
            //book.LoadFromFile("./wwwroot/data/trend/visinh/Raw/Raw data for of microbial results of room quality monitoring_  RD _2024.xlsx", ExcelVersion.Version2013);

            var worksheets = book.Worksheets.Count();
            var list_Result = new List<ResultModel>();
            for (var worksheetsIndex = 0; worksheetsIndex < worksheets; worksheetsIndex++)
            {
                Spire.Xls.Worksheet sheet = book.Worksheets[worksheetsIndex];
                var lastrow = sheet.LastDataRow;
                var lastcol = sheet.LastDataColumn;
                // nếu vẫn chưa gặp end thì vẫn lấy data
                Console.WriteLine(lastrow);
                var location_id = 1;
                var location = "";
                var stt = 0;
                for (int rowIndex = 9; rowIndex < lastrow; rowIndex++)
                {
                    // lấy row hiện tại
                    Console.WriteLine("rowIndex: {0} ", rowIndex);
                    var nowRow = sheet.Rows[rowIndex];
                    if (nowRow == null)
                        continue;

                    DateTime? date = nowRow.Cells[0] != null && nowRow.Cells[0].Value != "" ? nowRow.Cells[0].DateTimeValue : null;
                    if (date == null)
                        continue;
                    Console.WriteLine("date: {0} ", date);
                    var nowRowVitri = sheet.Rows[8];
                    for (int columnIndex = 2; columnIndex < lastcol; columnIndex++)
                    {
                        var code_vitri = nowRowVitri.Cells[columnIndex] != null ? nowRowVitri.Cells[columnIndex].Value : null;
                        if (code_vitri == null)
                            continue;
                        code_vitri = code_vitri.Trim();
                        Console.WriteLine("MS: {0} ", code_vitri);
                        var findPoint = _context.PointModel.Where(d => d.code == code_vitri).FirstOrDefault();
                        if (findPoint == null)
                        {
                            continue;
                        }

                        decimal? value = nowRow.Cells[columnIndex] != null && nowRow.Cells[columnIndex].Value != "NA" && nowRow.Cells[columnIndex].Value != "" ? (decimal)nowRow.Cells[columnIndex].NumberValue : null;
                        if (value != null)
                        {
                            var find_result = _context.ResultModel.Where(d => d.date == date && d.point_id == findPoint.id).FirstOrDefault();
                            if (find_result == null)
                            {
                                var result = new ResultModel()
                                {
                                    point_id = findPoint.id,
                                    value = value,
                                    target_id = findPoint.target_id,
                                    date = date,
                                    object_id = 2,
                                    created_at = DateTime.Now
                                };
                                list_Result.Add(result);
                                result.point = null;
                                _context.Add(result);
                            }
                            else
                            {
                                find_result.value = value;
                                _context.Update(find_result);
                            }
                            //_context.SaveChanges();
                        }
                    }


                    //var tansuat = 3;

                    //var chitieu_9 = nowRow.Cells[3] != null && nowRow.Cells[3].Value != "NA" && nowRow.Cells[3].Value != "" ? nowRow.Cells[3].Value : null;
                    //var chitieu_10 = nowRow.Cells[4] != null && nowRow.Cells[4].Value != "NA" && nowRow.Cells[4].Value != "" ? nowRow.Cells[4].Value : null;
                    //decimal? chitieu_11 = nowRow.Cells[5] != null && nowRow.Cells[5].Value != "NA" && nowRow.Cells[5].Value != "" ? (decimal)nowRow.Cells[5].NumberValue : null;
                    //decimal? chitieu_12 = nowRow.Cells[6] != null && nowRow.Cells[6].Value != "NA" && nowRow.Cells[6].Value != "" ? (decimal)nowRow.Cells[6].NumberValue : null;
                    //decimal? chitieu_13 = nowRow.Cells[7] != null && nowRow.Cells[7].Value != "NA" && nowRow.Cells[7].Value != "" ? (decimal)nowRow.Cells[7].NumberValue : null;
                    //var chitieu_14 = nowRow.Cells[8] != null && nowRow.Cells[8].Value != "NA" && nowRow.Cells[8].Value != "" ? nowRow.Cells[8].Value : null;




                    //EquipmentModel EquipmentModel = new EquipmentModel { code = code, name = name_vn, name_en = name_en, created_at = DateTime.Now };
                    //_context.Add(EquipmentModel);
                    //_context.SaveChanges();
                }
            }

            //_context.AddRange(list_Result);
            _context.SaveChanges();

            return Ok(list_Result);
        }
        public async Task<IActionResult> dataVisinhTPCN()
        {
            return Ok();
            // Khởi tạo workbook để đọc
            Spire.Xls.Workbook book = new Spire.Xls.Workbook();
            //book.LoadFromFile("./wwwroot/data/trend/visinh/Raw data for of microbial results of room quality monitoring_TPCN- GRADE D-2023.xlsx", ExcelVersion.Version2013);

            book.LoadFromFile("./wwwroot/data/trend/visinh/Raw/Raw data for of microbial results of room quality monitoring_TPCN- GRADE D-2024.xlsx", ExcelVersion.Version2013);

            var worksheets = book.Worksheets.Count();
            var list_Result = new List<ResultModel>();
            for (var worksheetsIndex = 0; worksheetsIndex < worksheets; worksheetsIndex++)
            {
                Spire.Xls.Worksheet sheet = book.Worksheets[worksheetsIndex];
                var lastrow = sheet.LastDataRow;
                var lastcol = sheet.LastDataColumn;
                // nếu vẫn chưa gặp end thì vẫn lấy data
                Console.WriteLine(lastrow);
                var location_id = 1;
                var location = "";
                var stt = 0;
                for (int rowIndex = 9; rowIndex < lastrow; rowIndex++)
                {
                    // lấy row hiện tại
                    Console.WriteLine("rowIndex: {0} ", rowIndex);
                    var nowRow = sheet.Rows[rowIndex];
                    if (nowRow == null)
                        continue;

                    DateTime? date = nowRow.Cells[0] != null && nowRow.Cells[0].Value != "" ? nowRow.Cells[0].DateTimeValue : null;
                    if (date == null)
                        continue;
                    Console.WriteLine("date: {0} ", date);
                    var nowRowVitri = sheet.Rows[8];
                    for (int columnIndex = 2; columnIndex < lastcol; columnIndex++)
                    {
                        var code_vitri = nowRowVitri.Cells[columnIndex] != null ? nowRowVitri.Cells[columnIndex].Value : null;
                        if (code_vitri == null)
                            continue;
                        code_vitri = code_vitri.Trim();
                        Console.WriteLine("MS: {0} ", code_vitri);
                        var findPoint = _context.PointModel.Where(d => d.code == code_vitri).FirstOrDefault();
                        if (findPoint == null)
                        {
                            continue;
                        }

                        decimal? value = nowRow.Cells[columnIndex] != null && nowRow.Cells[columnIndex].Value != "NA" && nowRow.Cells[columnIndex].Value != "" ? (decimal)nowRow.Cells[columnIndex].NumberValue : null;
                        if (value != null)
                        {
                            var find_result = _context.ResultModel.Where(d => d.date == date && d.point_id == findPoint.id).FirstOrDefault();
                            if (find_result == null)
                            {
                                var result = new ResultModel()
                                {
                                    point_id = findPoint.id,
                                    value = value,
                                    target_id = findPoint.target_id,
                                    date = date,
                                    created_at = DateTime.Now
                                };
                                list_Result.Add(result);
                                _context.Add(result);
                            }
                            else
                            {
                                find_result.value = value;
                                _context.Update(find_result);
                            }
                            //_context.SaveChanges();
                        }
                    }


                    //var tansuat = 3;

                    //var chitieu_9 = nowRow.Cells[3] != null && nowRow.Cells[3].Value != "NA" && nowRow.Cells[3].Value != "" ? nowRow.Cells[3].Value : null;
                    //var chitieu_10 = nowRow.Cells[4] != null && nowRow.Cells[4].Value != "NA" && nowRow.Cells[4].Value != "" ? nowRow.Cells[4].Value : null;
                    //decimal? chitieu_11 = nowRow.Cells[5] != null && nowRow.Cells[5].Value != "NA" && nowRow.Cells[5].Value != "" ? (decimal)nowRow.Cells[5].NumberValue : null;
                    //decimal? chitieu_12 = nowRow.Cells[6] != null && nowRow.Cells[6].Value != "NA" && nowRow.Cells[6].Value != "" ? (decimal)nowRow.Cells[6].NumberValue : null;
                    //decimal? chitieu_13 = nowRow.Cells[7] != null && nowRow.Cells[7].Value != "NA" && nowRow.Cells[7].Value != "" ? (decimal)nowRow.Cells[7].NumberValue : null;
                    //var chitieu_14 = nowRow.Cells[8] != null && nowRow.Cells[8].Value != "NA" && nowRow.Cells[8].Value != "" ? nowRow.Cells[8].Value : null;




                    //EquipmentModel EquipmentModel = new EquipmentModel { code = code, name = name_vn, name_en = name_en, created_at = DateTime.Now };
                    //_context.Add(EquipmentModel);
                    //_context.SaveChanges();
                }
            }

            //_context.AddRange(list_Result);
            _context.SaveChanges();

            return Ok(list_Result);
        }

        public async Task<IActionResult> dataVisinhQC()
        {
            //return Ok();
            // Khởi tạo workbook để đọc
            Spire.Xls.Workbook book = new Spire.Xls.Workbook();
            book.LoadFromFile("./wwwroot/data/trend/visinh/Raw/Raw data for of microbial results of room quality monitoring_QC- GRADE D.xlsx", ExcelVersion.Version2013);
            //book.LoadFromFile("./wwwroot/data/trend/visinh/Raw/Raw data for of microbial results of  equiment quality monitoring_QC- GRADE C.xlsx", ExcelVersion.Version2013);

            var worksheets = book.Worksheets.Count();
            var list_Result = new List<ResultModel>();
            for (var worksheetsIndex = 0; worksheetsIndex < worksheets; worksheetsIndex++)
            {
                Spire.Xls.Worksheet sheet = book.Worksheets[worksheetsIndex];
                var lastrow = sheet.LastDataRow;
                var lastcol = sheet.LastDataColumn;
                // nếu vẫn chưa gặp end thì vẫn lấy data
                Console.WriteLine(lastrow);
                var location_id = 1;
                var location = "";
                var stt = 0;
                for (int rowIndex = 8; rowIndex < lastrow; rowIndex++)
                {
                    // lấy row hiện tại
                    Console.WriteLine("rowIndex: {0} ", rowIndex);
                    var nowRow = sheet.Rows[rowIndex];
                    if (nowRow == null)
                        continue;
                    var code_vitri = nowRow.Cells[4] != null && nowRow.Cells[4].Value != "" ? nowRow.Cells[4].Value : null;
                    if (code_vitri == null)
                        continue;
                    code_vitri = code_vitri.Trim();
                    var findPoint = _context.PointModel.Where(d => d.code == code_vitri).FirstOrDefault();
                    if (findPoint == null)
                    {
                        continue;
                    }
                    Console.WriteLine("vitri: {0} ", code_vitri);
                    var nowRowDate = sheet.Rows[7];
                    for (int columnIndex = 5; columnIndex < lastcol; columnIndex++)
                    {

                        DateTime? date = nowRowDate.Cells[columnIndex] != null && nowRowDate.Cells[columnIndex].Value != "" ? nowRowDate.Cells[columnIndex].DateTimeValue : null;
                        if (date == null)
                            continue;

                        Console.WriteLine("date: {0} ", date);


                        decimal? value = nowRow.Cells[columnIndex] != null && nowRow.Cells[columnIndex].Value != "NA" && nowRow.Cells[columnIndex].Value != "" ? (decimal)nowRow.Cells[columnIndex].NumberValue : null;
                        if (value != null)
                        {
                            var find_result = _context.ResultModel.Where(d => d.date == date && d.point_id == findPoint.id).FirstOrDefault();
                            if (find_result == null)
                            {
                                var result = new ResultModel()
                                {
                                    point_id = findPoint.id,
                                    value = value,
                                    target_id = findPoint.target_id,
                                    date = date,
                                    created_at = DateTime.Now
                                };
                                list_Result.Add(result);
                                _context.Add(result);
                            }
                            else
                            {
                                find_result.value = value;
                                _context.Update(find_result);
                            }
                            //_context.SaveChanges();
                        }
                    }


                    //var tansuat = 3;

                    //var chitieu_9 = nowRow.Cells[3] != null && nowRow.Cells[3].Value != "NA" && nowRow.Cells[3].Value != "" ? nowRow.Cells[3].Value : null;
                    //var chitieu_10 = nowRow.Cells[4] != null && nowRow.Cells[4].Value != "NA" && nowRow.Cells[4].Value != "" ? nowRow.Cells[4].Value : null;
                    //decimal? chitieu_11 = nowRow.Cells[5] != null && nowRow.Cells[5].Value != "NA" && nowRow.Cells[5].Value != "" ? (decimal)nowRow.Cells[5].NumberValue : null;
                    //decimal? chitieu_12 = nowRow.Cells[6] != null && nowRow.Cells[6].Value != "NA" && nowRow.Cells[6].Value != "" ? (decimal)nowRow.Cells[6].NumberValue : null;
                    //decimal? chitieu_13 = nowRow.Cells[7] != null && nowRow.Cells[7].Value != "NA" && nowRow.Cells[7].Value != "" ? (decimal)nowRow.Cells[7].NumberValue : null;
                    //var chitieu_14 = nowRow.Cells[8] != null && nowRow.Cells[8].Value != "NA" && nowRow.Cells[8].Value != "" ? nowRow.Cells[8].Value : null;




                    //EquipmentModel EquipmentModel = new EquipmentModel { code = code, name = name_vn, name_en = name_en, created_at = DateTime.Now };
                    //_context.Add(EquipmentModel);
                    //_context.SaveChanges();
                }
            }

            //_context.AddRange(list_Result);
            _context.SaveChanges();

            return Ok(list_Result);
        }
        public async Task<JsonResult> colors()
        {
            var items = _context.PointModel.Where(_ => _.deleted_at == null).ToList();
            foreach (var item in items)
            {

                Color randomColor = GenerateRandomColor(item.code);
                item.color = ColorTranslator.ToHtml(randomColor);
            }
            _context.UpdateRange(items);
            _context.SaveChanges();
            return Json(new { success = true });
        }
        public static Color GenerateRandomColor(string input)
        {
            using (MD5 md5 = MD5.Create())
            {
                byte[] inputBytes = Encoding.UTF8.GetBytes(input);
                byte[] hashBytes = md5.ComputeHash(inputBytes);

                Color randomColor = Color.FromArgb(hashBytes[0], hashBytes[1], hashBytes[2]);

                return randomColor;
            }
        }
    }

}
