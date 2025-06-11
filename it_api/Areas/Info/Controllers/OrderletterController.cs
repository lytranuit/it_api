



using Info.Models;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.ComponentModel.DataAnnotations.Schema;
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
        private readonly ViewRender _view;
        public OrderletterController(NhansuContext context, AesOperation aes, TinhCong tinhcong, IConfiguration configuration, UserManager<UserModel> UserMgr, ViewRender view) : base(context, aes)
        {
            _configuration = configuration;
            UserManager = UserMgr;
            _tinhcong = tinhcong;
            _view = view;
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
        public async Task<JsonResult> Pheduyet(string id, string note)
        {
            System.Security.Claims.ClaimsPrincipal currentUser = this.User;
            var user_id = UserManager.GetUserId(currentUser);
            var Model = _context.OrderletterModel.Where(d => d.id == id).FirstOrDefault();



            if (user_id == Model.user_accept_id)
            {

                Model.date_accept = DateTime.Now;
                Model.user_accept_id = user_id;
                Model.note = note;
                Model.status_id = (int)OrderletterStatus.Duyet;
                _context.Update(Model);
                ///UPDATE CHẤM CÔNG
                var list_chamcong = _context.ChamcongModel.Where(d => d.orderletter_id == id).ToList();
                foreach (var d in list_chamcong)
                {
                    ///Copy
                    var details = new OrderletterDetailsModel()
                    {
                        id = d.id,
                        MANV = d.MANV,
                        NV_id = d.NV_id,
                        date = d.date,
                        shift_id = d.shift_id,
                        orderletter_id = d.orderletter_id,
                        value = d.value,
                        value_new = d.value_new,
                    };
                    _context.Add(details);

                    d.value = d.value_new;
                    d.orderletter_id = null;
                }
                _context.UpdateRange(list_chamcong);
                _context.SaveChanges();


                var user_created = _context.UserModel.Where(d => d.Id == Model.created_by).FirstOrDefault();
                var user_apply = _context.UserModel.Where(d => d.Id == Model.user_id).FirstOrDefault();
                var person_apply = _context.PersonnelModel.Where(d => d.EMAIL.ToLower() == user_apply.Email.ToLower()).FirstOrDefault();
                var user_related = new List<string>() { user_created.Email, user_apply.Email };
                //if (person_apply != null && person_apply.DIADIEM == "Hồ Chí Minh")
                //{
                //    user_related.Add("thu.ttn@astahealthcare.com");
                //}
                //else
                //{
                //    user_related.Add("hcns@astahealthcare.com");
                //}
                user_related = user_related.Distinct().ToList();

                var mail_string = string.Join(",", user_related);
                string Domain = (HttpContext.Request.IsHttps ? "https://" : "http://") + HttpContext.Request.Host.Value;
                var body = _view.Render("Emails/Orderletter_Success", new
                {
                    link_logo = Domain + "/images/clientlogo_astahealthcare.com_f1800.png",
                    link = _configuration["Application:Info:link"] + "OrderLetter/edit/" + Model.id
                });
                var email = new EmailModel
                {
                    email_to = mail_string,
                    subject = "[Thành công] " + Model.name,
                    body = body,
                    email_type = "Orderletter_Success",
                    status = 1
                };
                _context.Add(email);
                _context.SaveChanges();
            }
            else if (user_id == Model.user1_accept_id)
            {
                Model.date1_accept = DateTime.Now;
                Model.user1_accept_id = user_id;
                Model.note1 = note;
                Model.status1_id = (int)OrderletterStatus.Duyet;
                _context.Update(Model);

                ///Send mail thông báo đên người phê duyệt




                var details = _context.ChamcongModel.Where(d => d.orderletter_id != null && d.orderletter_id == Model.id).Include(d => d.shift).Select(d => new
                {
                    id = d.id,
                    shift = d.shift,
                    date = d.date,
                    value = d.value,
                    value_new = d.value_new
                }).OrderBy(d => d.date).ThenBy(d => d.shift).ToList();
                var user_accept_id = Model.user_accept_id;

                var user_pheduyet = _context.UserModel.Where(d => d.Id == user_accept_id).FirstOrDefault();
                var mail_string = user_pheduyet.Email;

                var user_apply = _context.UserModel.Where(d => d.Id == Model.user_id).FirstOrDefault();
                string Domain = (HttpContext.Request.IsHttps ? "https://" : "http://") + HttpContext.Request.Host.Value;
                var body = _view.Render("Emails/Orderletter", new
                {
                    link_logo = Domain + "/images/clientlogo_astahealthcare.com_f1800.png",
                    link = _configuration["Application:Info:link"] + "OrderLetter/edit/" + Model.id,
                    data = details,
                    user = user_apply.FullName
                });
                var email = new EmailModel
                {
                    email_to = mail_string,
                    subject = "[Đơn báo] " + Model.name,
                    body = body,
                    email_type = "Orderletter",
                    status = 1
                };
                _context.Add(email);
                _context.SaveChanges();

            }
            return Json(new { success = true, data = Model });
        }
        [HttpPost]
        public async Task<JsonResult> Khongduyet(string id, string note)
        {
            System.Security.Claims.ClaimsPrincipal currentUser = this.User;
            var user_id = UserManager.GetUserId(currentUser);
            var Model = _context.OrderletterModel.Where(d => d.id == id).FirstOrDefault();
            Model.status_id = (int)OrderletterStatus.Khong_duyet;
            if (Model.user1_accept_id != Model.user_accept_id && Model.user1_accept_id == user_id)
            {
                Model.date1_accept = DateTime.Now;
                Model.note1 = note;
                Model.status1_id = (int)OrderletterStatus.Khong_duyet;
            }
            else
            {
                Model.date_accept = DateTime.Now;
                Model.note = note;
            }
            ///UPDATE CHẤM CÔNG
            var list_chamcong = _context.ChamcongModel.Where(d => d.orderletter_id == id).ToList();
            foreach (var d in list_chamcong)
            {
                ///Copy
                var details = new OrderletterDetailsModel()
                {
                    id = d.id,
                    MANV = d.MANV,
                    NV_id = d.NV_id,
                    date = d.date,
                    shift_id = d.shift_id,
                    orderletter_id = d.orderletter_id,
                    value = d.value,
                    value_new = d.value_new,
                };
                _context.Add(details);

                d.value_new = d.value;
                d.orderletter_id = null;
            }
            _context.UpdateRange(list_chamcong);
            _context.Update(Model);
            _context.SaveChanges();

            var user_created = _context.UserModel.Where(d => d.Id == Model.created_by).FirstOrDefault();
            var user_apply = _context.UserModel.Where(d => d.Id == Model.user_id).FirstOrDefault();
            var user_related = new List<string>() { user_created.Email, user_apply.Email };
            user_related = user_related.Distinct().ToList();
            var mail_string = string.Join(",", user_related);

            string Domain = (HttpContext.Request.IsHttps ? "https://" : "http://") + HttpContext.Request.Host.Value;
            var body = _view.Render("Emails/Orderletter_Failed", new
            {
                link_logo = Domain + "/images/clientlogo_astahealthcare.com_f1800.png",
                link = _configuration["Application:Info:link"] + "OrderLetter/edit/" + Model.id,
                note = note
            });
            var email = new EmailModel
            {
                email_to = mail_string,
                subject = "[Thất bại] " + Model.name,
                body = body,
                email_type = "Orderletter_Failed",
                status = 1
            };
            _context.Add(email);
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

            ///Send mail thông báo đên người phê duyệt




            var details = _context.ChamcongModel.Where(d => d.orderletter_id != null && d.orderletter_id == OrderletterModel_old.id).Include(d => d.shift).Select(d => new
            {
                id = d.id,
                shift = d.shift,
                date = d.date,
                value = d.value,
                value_new = d.value_new
            }).OrderBy(d => d.date).ThenBy(d => d.shift).ToList();
            var user_accept_id = OrderletterModel_old.user_accept_id;
            if (OrderletterModel_old.user_accept_id != OrderletterModel_old.user1_accept_id)
            {
                user_accept_id = OrderletterModel_old.user1_accept_id;
            }
            var user_pheduyet = _context.UserModel.Where(d => d.Id == user_accept_id).FirstOrDefault();
            var mail_string = user_pheduyet.Email;

            var user_apply = _context.UserModel.Where(d => d.Id == OrderletterModel_old.user_id).FirstOrDefault();
            string Domain = (HttpContext.Request.IsHttps ? "https://" : "http://") + HttpContext.Request.Host.Value;
            var body = _view.Render("Emails/Orderletter", new
            {
                link_logo = Domain + "/images/clientlogo_astahealthcare.com_f1800.png",
                link = _configuration["Application:Info:link"] + "OrderLetter/edit/" + OrderletterModel_old.id,
                data = details,
                user = user_apply.FullName
            });
            var email = new EmailModel
            {
                email_to = mail_string,
                subject = "[Đơn báo] " + OrderletterModel_old.name,
                body = body,
                email_type = "Orderletter",
                status = 1
            };
            _context.Add(email);
            _context.SaveChanges();
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
            var created_by = Request.Form["filters[created_by]"].FirstOrDefault();
            var user_id1 = Request.Form["filters[user_id]"].FirstOrDefault();
            var user_accept_id = Request.Form["filters[user_accept_id]"].FirstOrDefault();
            var user1_accept_id = Request.Form["filters[user1_accept_id]"].FirstOrDefault();
            var status = Request.Form["filters[status]"].FirstOrDefault();
            var status_id = status != null ? Convert.ToInt32(status) : 0;
            var filter_phep = Request.Form["filters[filter_phep]"].FirstOrDefault();
            //var filter_phep_bool = filter_phep != null ? Convert.ToBoolean(filter_phep) : false;
            //var tenhh = Request.Form["filters[tenhh]"].FirstOrDefault();
            int skip = start != null ? Convert.ToInt32(start) : 0;

            /// CHECK PHAN QUYEN
            var user = await UserManager.GetUserAsync(currentUser);
            var email = user.Email;
            var is_admin = await UserManager.IsInRoleAsync(user, "Administrator");
            var is_manager = await UserManager.IsInRoleAsync(user, "Manager HR");
            var is_hr = await UserManager.IsInRoleAsync(user, "HR");
            var customerData = _context.OrderletterModel.Where(d => d.deleted_at == null);
            if (is_admin || is_hr)
            {

            }
            else if (is_manager)
            {
                var person = _context.PersonnelModel.Where(d => d.EMAIL == email).FirstOrDefault();
                if (person != null)
                {
                    var maphong = person.MAPHONG;
                    var personinPhong = _context.PersonnelModel.Where(d => d.MAPHONG == maphong).Select(d => d.EMAIL).ToList();
                    if (email == "thao.pdp@astahealthcare.com")
                    {
                        personinPhong = _context.PersonnelModel.Where(d => d.MAPHONG == maphong || d.MAPHONG == "22").Select(d => d.EMAIL).ToList();
                    }

                    var userinPhong = _context.UserModel.Where(d => personinPhong.Contains(d.Email)).Select(d => d.Id).ToList();
                    customerData = customerData.Where(d => userinPhong.Contains(d.created_by));
                }
            }
            else
            {
                customerData = customerData.Where(d => d.created_by == user_id || d.user_accept_id == user_id || d.user1_accept_id == user_id);
            }

            if (type == "0")
            {
                customerData = customerData.Where(d => d.created_by == user_id);
            }
            else if (type == "1")
            {
                customerData = customerData.Where(d => (d.user_accept_id == user_id || d.user_accept_id == user_id) && d.status_id == 1);
            }
            else if (type == "2")
            {
                var orderletter = _context.OrderletterDetailsModel.Where(d => d.value_new != "X").Select(d => d.orderletter_id).ToList();
                customerData = customerData.Where(d => orderletter.Contains(d.id));
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
            if (created_by != null)
            {
                customerData = customerData.Where(d => d.created_by == created_by);
            }
            if (user_id1 != null)
            {
                customerData = customerData.Where(d => d.user_id == user_id1);
            }
            if (user_accept_id != null)
            {
                customerData = customerData.Where(d => d.user_accept_id == user_accept_id);
            }
            if (user1_accept_id != null)
            {
                customerData = customerData.Where(d => d.user1_accept_id == user1_accept_id && d.user_accept_id != user1_accept_id);
            }

            if (status_id != 0)
            {
                customerData = customerData.Where(d => d.status_id == status_id);
            }
            //if (filter_phep_bool)
            //{
            //    var orderletter = _context.OrderletterDetailsModel.Where(d => d.value_new != "X").Select(d => d.orderletter_id).ToList();
            //    customerData = customerData.Where(d => orderletter.Contains(d.id));
            //}
            int recordsFiltered = customerData.Count();
            var datapost = customerData.OrderByDescending(d => d.created_at).Skip(skip).Take(pageSize).Include(d => d.user_created_by).Include(d => d.user).Include(d => d.user_accept).Include(d => d.user1_accept).ToList();
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
            //var list_nv = datapost.Select(d => d.MANV).ToList();

            var date_lock = _context.OptionModel.Where(d => d.key == "date_lock").FirstOrDefault();
            var jsonData = new { data = data, date_lock = date_lock != null ? date_lock.date_value : null };
            return Json(jsonData, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }

        public JsonResult Get(string id)
        {
            var data = _context.OrderletterModel.Where(d => d.id == id).Include(d => d.user_created_by).Include(d => d.user_accept).Include(d => d.user1_accept).FirstOrDefault();
            return Json(data, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }
        public JsonResult getListChamcong(string id, string user_id)
        {
            var user = _context.UserModel.Where(d => d.Id == user_id).FirstOrDefault();
            var customerData = _context.PersonnelModel.Where(d => d.EMAIL.ToLower() == user.Email.ToLower()).FirstOrDefault();
            var OrderLetterModel = _context.OrderletterModel.Where(d => d.id == id).FirstOrDefault();
            var ChamcongModel = _context.ChamcongModel.Where(d => customerData.id == d.NV_id && d.orderletter_id != null && d.orderletter_id == id).Include(d => d.shift).Select(d => new ChamcongModel
            {
                id = d.id,
                shift = d.shift,
                shift_id = d.shift_id,
                NV_id = d.NV_id,
                MANV = d.MANV,
                date = d.date,
                value = d.value,
                value_new = d.value_new,
                orderletter_id = d.orderletter_id,
                is_duyet = d.is_duyet,
            }).ToList();
            ///KHI ĐÃ PHÊ DUYỆT
            if (OrderLetterModel.status_id != (int)OrderletterStatus.New)
            {
                ChamcongModel = _context.OrderletterDetailsModel.Where(d => customerData.id == d.NV_id && d.orderletter_id != null && d.orderletter_id == id).Include(d => d.shift).Select(d => new ChamcongModel
                {
                    id = d.id,
                    shift = d.shift,
                    shift_id = d.shift_id,
                    NV_id = d.NV_id,
                    MANV = d.MANV,
                    date = d.date,
                    value = d.value,
                    value_new = d.value_new,
                    orderletter_id = d.orderletter_id,
                    is_duyet = true,
                }).ToList();
            }

            return Json(ChamcongModel, new System.Text.Json.JsonSerializerOptions()
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
