



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
using System.Data;
using System.Text.Json.Serialization;
using Vue.Data;
using Vue.Models;
using Vue.Services;

namespace it_template.Areas.Info.Controllers
{

	public class OrderletterController : BaseController
	{
		private readonly IConfiguration _configuration;
		private UserManager<UserModel> UserManager;
		private readonly TinhCong _tinhcong;
		public OrderletterController(NhansuContext context, AesOperation aes, TinhCong tinhcong, IConfiguration configuration, UserManager<UserModel> UserMgr) : base(context, aes)
		{
			_configuration = configuration;
			UserManager = UserMgr;
			_tinhcong = tinhcong;

		}

		[HttpPost]
		public async Task<JsonResult> Delete(string id)
		{
			var Model = _context.OrderletterModel.Where(d => d.id == id).FirstOrDefault();
			Model.deleted_at = DateTime.Now;
			_context.Update(Model);
			var chamcong = _context.ChamcongModel.Where(d => d.orderletter_id == id).ToList();
			foreach (var d in chamcong)
			{
				d.orderletter_id = null;
				d.is_duyet = null;
				d.value_new = d.value;
			}
			_context.UpdateRange(chamcong);
			_context.SaveChanges();
			return Json(new { success = true, data = Model });
		}
		[HttpPost]
		public async Task<JsonResult> Pheduyet(string id)
		{
			System.Security.Claims.ClaimsPrincipal currentUser = this.User;
			var user_id = UserManager.GetUserId(currentUser);
			var Model = _context.OrderletterModel.Where(d => d.id == id).FirstOrDefault();
			Model.date_accept = DateTime.Now;
			Model.user_accept_id = user_id;
			Model.status_id = (int)OrderletterStatus.Duyet;
			_context.Update(Model);

			///UPDATE CHẤM CÔNG
			var list_chamcong = _context.ChamcongModel.Where(d => d.orderletter_id == id).ToList();
			foreach (var d in list_chamcong)
			{
				d.value = d.value_new;
			}
			_context.UpdateRange(list_chamcong);
			_context.SaveChanges();
			return Json(new { success = true, data = Model });
		}
		[HttpPost]
		public async Task<JsonResult> Khongduyet(string id, string note)
		{
			System.Security.Claims.ClaimsPrincipal currentUser = this.User;
			var user_id = UserManager.GetUserId(currentUser);
			var Model = _context.OrderletterModel.Where(d => d.id == id).FirstOrDefault();
			Model.date_accept = DateTime.Now;
			Model.status_id = (int)OrderletterStatus.Khong_duyet;
			Model.user_accept_id = user_id;
			Model.note = note;
			_context.Update(Model);
			_context.SaveChanges();
			return Json(new { success = true, data = Model });
		}
		[HttpPost]
		public async Task<JsonResult> Save(OrderletterModel OrderletterModel, List<ChamcongModel> list)
		{
			System.Security.Claims.ClaimsPrincipal currentUser = this.User;
			var user_id = UserManager.GetUserId(currentUser);
			var user = await UserManager.GetUserAsync(currentUser);
			OrderletterModel? OrderletterModel_old;
			if (OrderletterModel.id == null)
			{
				OrderletterModel.id = Guid.NewGuid().ToString();
				OrderletterModel.created_at = DateTime.Now;

				OrderletterModel.created_by = user_id;
				OrderletterModel.status_id = (int)OrderletterStatus.New;


				_context.OrderletterModel.Add(OrderletterModel);

				_context.SaveChanges();

				OrderletterModel_old = OrderletterModel;

			}
			else
			{

				OrderletterModel_old = _context.OrderletterModel.Where(d => d.id == OrderletterModel.id).FirstOrDefault();
				CopyValues<OrderletterModel>(OrderletterModel_old, OrderletterModel);
				OrderletterModel_old.updated_at = DateTime.Now;

				_context.Update(OrderletterModel_old);
				_context.SaveChanges();
			}
			/////
			if (list != null && list.Count() > 0)
			{
				foreach (var item in list)
				{
					if (item.value != item.value_new && item.orderletter_id == null)
					{
						item.orderletter_id = OrderletterModel_old.id;
					}
					else if (item.orderletter_id != null && item.value == item.value_new)
					{
						item.orderletter_id = null;
					}

					if (item.id == null || item.id == 0)
					{
						_context.Add(item);
					}
					else
					{
						_context.Update(item);
					}
				}
				_context.SaveChanges();
			}

			return Json(new { success = true, data = OrderletterModel_old });
		}


		[HttpPost]
		public async Task<JsonResult> Table()
		{
			System.Security.Claims.ClaimsPrincipal currentUser = this.User;
			var user_id = UserManager.GetUserId(currentUser);
			var type = Request.Form["type"].FirstOrDefault();
			var draw = Request.Form["draw"].FirstOrDefault();
			var start = Request.Form["start"].FirstOrDefault();
			var length = Request.Form["length"].FirstOrDefault();
			int pageSize = length != null ? Convert.ToInt32(length) : 0;
			var name = Request.Form["filters[name]"].FirstOrDefault();
			var id = Request.Form["filters[id]"].FirstOrDefault();
			//var tenhh = Request.Form["filters[tenhh]"].FirstOrDefault();
			int skip = start != null ? Convert.ToInt32(start) : 0;
			var customerData = _context.OrderletterModel.Where(d => d.deleted_at == null && (d.created_by == user_id || d.user_accept_id == user_id));
			if (type == "0")
			{
				customerData = customerData.Where(d => d.created_by == user_id);
			}
			else if (type == "1")
			{
				customerData = customerData.Where(d => d.user_accept_id == user_id && d.status_id == 1);
			}
			int recordsTotal = customerData.Count();

			if (name != null && name != "")
			{
				customerData = customerData.Where(d => d.name.Contains(name));
			}
			if (id != null)
			{
				customerData = customerData.Where(d => d.id == id);
			}

			int recordsFiltered = customerData.Count();
			var datapost = customerData.OrderByDescending(d => d.created_at).Skip(skip).Take(pageSize).Include(d => d.user_created_by).ToList();
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
		public async Task<JsonResult> TableWorking()
		{   /// CHECK PHAN QUYEN
			//System.Security.Claims.ClaimsPrincipal currentUser = this.User;
			//var user = await UserManager.GetUserAsync(currentUser);
			//var user_id = user.Id;


			var id = Request.Form["id"].FirstOrDefault();
			var user_id = Request.Form["user_id"].FirstOrDefault();
			var user = _context.UserModel.Where(d => d.Id == user_id).FirstOrDefault();
			var customerData = _context.PersonnelModel.Where(d => d.EMAIL.ToLower() == user.Email.ToLower());
			int recordsTotal = customerData.Count();
			var date_from_string = Request.Form["filters[date_from]"].FirstOrDefault();
			var date_to_string = Request.Form["filters[date_to]"].FirstOrDefault();
			DateTime date_from = date_from_string != null ? date_from = DateTime.ParseExact(date_from_string, "yyyy-MM-dd",
										   System.Globalization.CultureInfo.InvariantCulture) : date_from = DateTime.Now;
			DateTime date_to = date_to_string != null ? date_to = DateTime.ParseExact(date_to_string, "yyyy-MM-dd",
										   System.Globalization.CultureInfo.InvariantCulture) : date_to = DateTime.Now;

			int recordsFiltered = customerData.Count();
			var datapost = customerData.ToList();

			var data = _tinhcong.cong(datapost, date_from, date_to);
			var list_nv = datapost.Select(d => d.MANV).ToList();
			var ChamcongModel = _context.ChamcongModel.Where(d => list_nv.Contains(d.MANV) && d.orderletter_id != null && d.orderletter_id == id).Include(d => d.shift).ToList();

			var date_lock = _context.OptionModel.Where(d => d.key == "date_lock").FirstOrDefault();
			var jsonData = new { data = data, list_chamcong = ChamcongModel, date_lock = date_lock != null ? date_lock.date_value : null };
			return Json(jsonData, new System.Text.Json.JsonSerializerOptions()
			{
				DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
			});
		}

		public JsonResult Get(string id)
		{
			var data = _context.OrderletterModel.Where(d => d.id == id).Include(d => d.user_created_by).Include(d => d.user_accept).FirstOrDefault();
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
	}

}
