

using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json.Linq;
using Spire.Xls;
using System.Collections;
using Vue.Data;

namespace it_template.Areas.V1.Controllers
{

	public class HistoryController : BaseController
	{
		private readonly IConfiguration _configuration;
		public HistoryController(ItContext context, IConfiguration configuration) : base(context)
		{
			_configuration = configuration;
		}

		[HttpPost]
		public async Task<JsonResult> Table()
		{
			var draw = Request.Form["draw"].FirstOrDefault();
			var start = Request.Form["start"].FirstOrDefault();
			var length = Request.Form["length"].FirstOrDefault();
			var searchValue = Request.Form["search[value]"].FirstOrDefault();
			var filter_user = Request.Form["filters[user]"].FirstOrDefault();
			var type = Request.Form["filters[type]"].FirstOrDefault();
			var tableName = Request.Form["filters[tableName]"].FirstOrDefault();
			var primaryKey = Request.Form["filters[primaryKey]"].FirstOrDefault();
			var oldValues = Request.Form["filters[oldValues]"].FirstOrDefault();
			var newValues = Request.Form["filters[newValues]"].FirstOrDefault();
			var description = Request.Form["filters[description]"].FirstOrDefault();
			var datetime = Request.Form["datetime[]"].ToList();
			int pageSize = length != null ? Convert.ToInt32(length) : 0;
			int skip = start != null ? Convert.ToInt32(start) : 0;
			var customerData = _context.AuditTrailsModel.Where(d => d.UserId != "5a375cd2-1908-4784-9b7b-d470e2d63376");
			if (datetime.Count > 1)
			{
				DateTime start_date = DateTime.Parse(datetime[0].ToString());
				DateTime end_date = DateTime.Parse(datetime[1].ToString());

				customerData = customerData.Where(m => m.DateTime.Date >= start_date.Date && m.DateTime.Date <= end_date.Date);
			}


			if (filter_user != null && filter_user != "")
			{
				var user_list = _context.UserModel.Where(d => d.FullName.Contains(filter_user)).Select(d => d.Id).ToList();
				customerData = customerData.Where(d => user_list.Contains(d.UserId));
			}

			if (oldValues != null && oldValues != "")
			{
				customerData = customerData.Where(d => d.OldValues.Contains(oldValues));
			}
			if (newValues != null && newValues != "")
			{
				customerData = customerData.Where(d => d.NewValues.Contains(newValues));
			}
			if (primaryKey != null && primaryKey != "")
			{
				customerData = customerData.Where(d => d.PrimaryKey.Contains(primaryKey));
			}
			if (tableName != null && tableName != "")
			{
				customerData = customerData.Where(d => d.TableName.Contains(tableName));
			}
			if (type != null && type != "")
			{
				customerData = customerData.Where(d => d.Type.Contains(type));
			}
			if (description != null && description != "")
			{
				customerData = customerData.Where(d => d.description.Contains(description));
			}
			int recordsTotal = customerData.Count();
			int recordsFiltered = customerData.Count();
			var datapost = customerData.Include(d => d.user).OrderByDescending(d => d.Id).Skip(skip).Take(pageSize).ToList();
			var data = new ArrayList();
			foreach (var record in datapost)
			{
				var user = record.user;
				var user_name = "";
				if (user != null)
				{
					user_name = user.FullName;

				}
				var data1 = new
				{
					id = record.Id,
					user = user_name,
					datetime = record.DateTime.ToString("yyyy/MM/dd HH:mm:ss"),
					type = record.Type,
					tableName = record.TableName,
					description = record.description,
					oldValues = $"<div style='white-space: pre-wrap;'>{record.OldValues}</div>",
					newValues = $"<div style='white-space: pre-wrap;'>{record.NewValues}</div>",
					primaryKey = record.PrimaryKey,
				};
				data.Add(data1);
			}
			var jsonData = new { draw = draw, recordsFiltered = recordsFiltered, recordsTotal = recordsTotal, data = data };
			return Json(jsonData);
		}

		public async Task<JsonResult> ExportExcel()
		{
			var draw = Request.Form["draw"].FirstOrDefault();
			var start = Request.Form["start"].FirstOrDefault();
			var length = Request.Form["length"].FirstOrDefault();
			var searchValue = Request.Form["search[value]"].FirstOrDefault();
			var filter_user = Request.Form["filters[user]"].FirstOrDefault();
			var description = Request.Form["filters[description]"].FirstOrDefault();
			var datetime = Request.Form["datetime[]"].ToList();
			int pageSize = length != null ? Convert.ToInt32(length) : 0;
			int skip = start != null ? Convert.ToInt32(start) : 0;
			var customerData = _context.AuditTrailsModel.Where(d => d.UserId != "5a375cd2-1908-4784-9b7b-d470e2d63376");
			if (datetime.Count > 1)
			{
				DateTime start_date = DateTime.Parse(datetime[0].ToString());
				DateTime end_date = DateTime.Parse(datetime[1].ToString());

				customerData = customerData.Where(m => m.DateTime.Date >= start_date.Date && m.DateTime.Date <= end_date.Date);
			}


			if (filter_user != null && filter_user != "")
			{
				var user_list = _context.UserModel.Where(d => d.FullName.Contains(filter_user)).Select(d => d.Id).ToList();
				customerData = customerData.Where(d => user_list.Contains(d.UserId));
			}
			if (description != null && description != "")
			{
				customerData = customerData.Where(d => d.description.Contains(description));
			}
			int recordsTotal = customerData.Count();
			int recordsFiltered = customerData.Count();
			var datapost = customerData.Include(d => d.user).OrderByDescending(d => d.Id).ToList();
			var data = new ArrayList();
			var viewPath = "./wwwroot/report/excel/AuditTrails.xlsx";
			var documentPath = "/tmp/" + DateTime.Now.ToFileTime() + ".xlsx";
			Workbook workbook = new Workbook();
			workbook.LoadFromFile(viewPath);
			Worksheet sheet = workbook.Worksheets[0];

			ExcelFont fontItalic1 = workbook.CreateFont();
			fontItalic1.IsItalic = true;
			fontItalic1.Size = 10;
			fontItalic1.FontName = "Arial";
			fontItalic1.IsBold = true;
			var row_c = 1;
			var cur_r = 1;
			foreach (var record in datapost)
			{
				sheet.InsertRow(cur_r + row_c + 1);
				sheet.Copy(sheet.Range["A" + (cur_r + row_c) + ":J" + (cur_r + row_c)], sheet.Range["A" + (cur_r + row_c + 1) + ":J" + (cur_r + row_c + 1)], true);
				var row = sheet.Rows[cur_r + row_c - 1];
				row.Cells[0].Value = record.Id.ToString();
				row.Cells[1].Value = record.user != null ? record.user.FullName : "";
				row.Cells[2].Value = record.DateTime.ToString("yyyy/MM/dd HH:mm:ss");
				row.Cells[3].Value = record.Type;
				row.Cells[4].Value = record.description;
				row.Cells[5].Value = record.TableName;
				row.Cells[6].Value = record.PrimaryKey;
				row.Cells[7].Value = record.OldValues;
				row.Cells[8].Value = record.NewValues;
			}
			//AutoFit column width and row height
			sheet.AllocatedRange.AutoFitColumns();

			sheet.AllocatedRange.AutoFitRows();

			workbook.SaveToFile("./wwwroot" + documentPath, ExcelVersion.Version2013);


			string Domain = (HttpContext.Request.IsHttps ? "https://" : "http://") + HttpContext.Request.Host.Value;
			var jsonData = new { success = true, link = Domain + documentPath };
			return Json(jsonData);
		}
	}
}
