


using Info.Models;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using PH.WorkingDaysAndTimeUtility.Configuration;
using System.Collections;
using System.Data;
using System.Text.Json.Serialization;
using Vue.Data;
using Vue.Models;
using Vue.Services;
using Spire.Xls;
using Microsoft.EntityFrameworkCore;
using Spire.Xls.Collections;

namespace it_template.Areas.Info.Controllers
{

    public class WorkingController : BaseController
    {
        private readonly IConfiguration _configuration;
        private UserManager<UserModel> UserManager;
        private readonly TinhCong _tinhcong;
        public WorkingController(NhansuContext context, AesOperation aes, TinhCong tinhcong, IConfiguration configuration, UserManager<UserModel> UserMgr) : base(context, aes)
        {
            _configuration = configuration;
            UserManager = UserMgr;
            _tinhcong = tinhcong;
        }

        [HttpPost]
        public async Task<JsonResult> SaveChamcong(List<ChamcongModel> list)
        {

            if (list != null && list.Count() > 0)
            {
                foreach (var item in list)
                {
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
            return Json(new { success = true });
        }
        public async Task<JsonResult> holidays(int year)
        {
            var data = _context.HolidayModel.Where(d => d.date.Value.Year == year).ToList();
            return Json(data, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }
        [HttpPost]
        public async Task<JsonResult> SaveDateLock(DateTime date_lock)
        {
            var option = _context.OptionModel.Where(d => d.key == "date_lock").FirstOrDefault();
            if (option == null)
            {
                option = new OptionModel()
                {
                    key = "date_lock",
                    date_value = date_lock
                };
                _context.OptionModel.Add(option);
                _context.SaveChanges();
            }
            else
            {
                var date_from = option.date_value.Value;
                var date_to = date_lock;
                option.date_value = date_lock;
                _context.OptionModel.Update(option);
                _context.SaveChanges();
                if (date_to > date_from)
                {
                    var datapost = _context.PersonnelModel.Where(d => d.NGAYNGHIVIEC == null).ToList();
                    ////
                    var data = _tinhcong.cong(datapost, date_from, date_to, true);

                }


            }


            return Json(new { success = true });
        }
        [HttpPost]
        public async Task<JsonResult> SaveHolidays(List<HolidayModel> list_add, List<string> list_remove)
        {
            if (list_remove != null && list_remove.Count() > 0)
            {
                var list = _context.HolidayModel.Where(d => list_remove.Contains(d.id)).ToList();
                _context.RemoveRange(list);
                _context.SaveChanges();
            }
            if (list_add != null && list_add.Count() > 0)
            {
                foreach (var item in list_add)
                {
                    item.id = Guid.NewGuid().ToString();

                }
                _context.AddRange(list_add);
                _context.SaveChanges();
            }
            return Json(new { success = true });
        }

        [HttpPost]
        public async Task<JsonResult> Table()
        {

            var draw = Request.Form["draw"].FirstOrDefault();
            var start = Request.Form["start"].FirstOrDefault();
            var length = Request.Form["length"].FirstOrDefault();
            int pageSize = length != null ? Convert.ToInt32(length) : 0;
            var search = Request.Form["filters[search]"].FirstOrDefault();
            var id = Request.Form["filters[id]"].FirstOrDefault();
            var department = Request.Form["filters[department]"].FirstOrDefault();
            var date_from_string = Request.Form["filters[date_from]"].FirstOrDefault();
            var date_to_string = Request.Form["filters[date_to]"].FirstOrDefault();
            DateTime date_from = date_from_string != null ? date_from = DateTime.ParseExact(date_from_string, "yyyy-MM-dd",
                                           System.Globalization.CultureInfo.InvariantCulture) : date_from = DateTime.Now;
            DateTime date_to = date_to_string != null ? date_to = DateTime.ParseExact(date_to_string, "yyyy-MM-dd",
                                           System.Globalization.CultureInfo.InvariantCulture) : date_to = DateTime.Now;
            //TimeSpan start_time = new TimeSpan(10, 30, 0);
            //TimeSpan end_time = new TimeSpan(14, 0, 0);
            //var tenhh = Request.Form["filters[tenhh]"].FirstOrDefault();
            int skip = start != null ? Convert.ToInt32(start) : 0;
            var customerData = _context.PersonnelModel.Where(d => (d.NGAYNGHIVIEC == null || d.NGAYNGHIVIEC > date_from));


            /// CHECK PHAN QUYEN
            System.Security.Claims.ClaimsPrincipal currentUser = this.User;
            var user = await UserManager.GetUserAsync(currentUser);
            var user_id = user.Id;
            var email = user.Email;
            var is_admin = await UserManager.IsInRoleAsync(user, "Administrator");
            var is_manager = await UserManager.IsInRoleAsync(user, "Manager HR");
            var is_hr = await UserManager.IsInRoleAsync(user, "HR");

            if (is_manager)
            {
                var person = _context.PersonnelModel.Where(d => d.EMAIL == email).FirstOrDefault();
                if (person != null)
                {
                    var maphong = person.MAPHONG;
                    if (email == "thao.pdp@astahealthcare.com")
                    {
                        customerData = customerData.Where(d => d.MAPHONG == maphong || d.MAPHONG == "22");
                    }
                    else
                    {
                        customerData = customerData.Where(d => d.MAPHONG == maphong);
                    }
                }
                else
                {
                    customerData = customerData.Where(d => 0 == 1);
                }
            }
            else if (!is_admin && !is_hr)
            {
                customerData = customerData.Where(d => d.EMAIL == email);
            }




            int recordsTotal = customerData.Count();

            if (search != null && search != "")
            {
                customerData = customerData.Where(d => d.MANV.Contains(search) || d.HOVATEN.Contains(search));
            }
            if (department != null)
            {
                customerData = customerData.Where(d => d.MAPHONG == department);
            }
            if (id != null)
            {
                customerData = customerData.Where(d => d.id == id);
            }

            int recordsFiltered = customerData.Count();
            var datapost = customerData.OrderByDescending(d => d.NGAYNHANVIEC).ThenBy(d => d.MANV).Skip(skip).Take(pageSize).ToList();


            var data = _tinhcong.cong(datapost, date_from, date_to);
            var list_nv = datapost.Select(d => d.MANV).ToList();
            var ChamcongModel = _context.ChamcongModel.Where(d => d.date.Value.Date >= date_from && d.date.Value.Date <= date_to && list_nv.Contains(d.MANV)).ToList();

            var date_lock = _context.OptionModel.Where(d => d.key == "date_lock").FirstOrDefault();
            var jsonData = new { draw = draw, recordsFiltered = recordsFiltered, recordsTotal = recordsTotal, data = data, list_chamcong = ChamcongModel, date_lock = date_lock != null ? date_lock.date_value : null };
            return Json(jsonData, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }



        [HttpPost]
        public async Task<JsonResult> excel()
        {
            //try
            //{

            ////Lấy 
            var draw = Request.Form["draw"].FirstOrDefault();
            var start = Request.Form["start"].FirstOrDefault();
            var length = Request.Form["length"].FirstOrDefault();
            int pageSize = length != null ? Convert.ToInt32(length) : 0;
            var search = Request.Form["filters[search]"].FirstOrDefault();
            var id = Request.Form["filters[id]"].FirstOrDefault();
            var type = Request.Form["filters[type]"].FirstOrDefault();
            var department = Request.Form["filters[department]"].FirstOrDefault();
            var date_from_string = Request.Form["filters[date_from]"].FirstOrDefault();
            var date_to_string = Request.Form["filters[date_to]"].FirstOrDefault();
            DateTime date_from = date_from_string != null ? date_from = DateTime.ParseExact(date_from_string, "yyyy-MM-dd",
                                           System.Globalization.CultureInfo.InvariantCulture) : date_from = DateTime.Now;
            DateTime date_to = date_to_string != null ? date_to = DateTime.ParseExact(date_to_string, "yyyy-MM-dd",
                                           System.Globalization.CultureInfo.InvariantCulture) : date_to = DateTime.Now;
            TimeSpan start_time = new TimeSpan(10, 30, 0);
            TimeSpan end_time = new TimeSpan(14, 0, 0);
            var customerData = _context.PersonnelModel.Where(d => (d.NGAYNGHIVIEC == null || d.NGAYNGHIVIEC > date_from));

            /// CHECK PHAN QUYEN
            System.Security.Claims.ClaimsPrincipal currentUser = this.User;
            var user = await UserManager.GetUserAsync(currentUser);
            var user_id = user.Id;
            var email = user.Email;
            var is_admin = await UserManager.IsInRoleAsync(user, "Administrator");
            var is_manager = await UserManager.IsInRoleAsync(user, "Manager HR");
            var is_hr = await UserManager.IsInRoleAsync(user, "HR");

            if (is_manager)
            {
                var person = _context.PersonnelModel.Where(d => d.EMAIL == email).FirstOrDefault();
                if (person != null)
                {
                    var maphong = person.MAPHONG;
                    if (email == "thao.pdp@astahealthcare.com")
                    {
                        customerData = customerData.Where(d => d.MAPHONG == maphong || d.MAPHONG == "22");
                    }
                    else
                    {
                        customerData = customerData.Where(d => d.MAPHONG == maphong);
                    }
                }
                else
                {
                    customerData = customerData.Where(d => 0 == 1);
                }
            }
            else if (!is_admin && !is_hr)
            {
                customerData = customerData.Where(d => d.EMAIL == email);
            }

            if (department != null)
            {
                customerData = customerData.Where(d => d.MAPHONG == department);
            }
            if (search != null && search != "")
            {
                customerData = customerData.Where(d => d.MANV.Contains(search) || d.HOVATEN.Contains(search));
            }
            if (id != null)
            {
                customerData = customerData.Where(d => d.id == id);
            }

            int recordsFiltered = customerData.Count();
            var datapost_all = customerData.OrderBy(d => d.MANV).ToList();
            var data = new ArrayList();

            var list_nv = datapost_all.Select(d => d.MANV).ToList();
            var ChamcongModel = _context.ChamcongModel.Where(d => d.date.Value.Date >= date_from && d.date.Value.Date <= date_to && list_nv.Contains(d.MANV)).ToList();
            var list_hik = _context.HikModel.Where(d => d.date.Value.Date >= date_from && d.date.Value.Date <= date_to && d.device != "A.CT.1").OrderBy(d => d.date).ThenBy(d => d.time).ToList();
            var list_holidays = _context.HolidayModel.Where(d => d.date.Value.Date >= date_from && d.date.Value.Date <= date_to).Select(d => d.date).ToList();


            var user_shift_vs = _context.ShiftUserModel.Include(d => d.shift).Where(d => d.shift.deleted_at == null && (d.shift.code == "S-T7" || d.shift.code == "C-T7")).Select(d => d.person_id).Distinct().ToList();
            var user_shift_full = _context.ShiftUserModel.Include(d => d.shift).Where(d => d.shift.deleted_at == null && (d.shift.code == "F")).Select(d => d.person_id).Distinct().ToList();
            var user_shift = _context.ShiftUserModel.Include(d => d.shift).Where(d => d.shift.deleted_at == null && (d.shift.code == "S" || d.shift.code == "C")).Select(d => d.person_id).Distinct().ToList();



            ///EXCEL
            /// 
            var viewPath = "wwwroot/report/excel/Bảng chấm công.xlsx";
            var documentPath = "/tmp/" + DateTime.Now.ToFileTimeUtc() + ".xlsx";
            string Domain = (HttpContext.Request.IsHttps ? "https://" : "http://") + HttpContext.Request.Host.Value;
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(viewPath);


            ////Nhân viên chính thức / thử việc / học việc

            var datapost = datapost_all.Where(d => d.LOAIHD != "DV" && user_shift.Contains(d.id)).OrderByDescending(d => d.NGAYNHANVIEC).ToList();
            if (datapost.Count > 0)
            {
                Worksheet sheet = workbook.Worksheets[0];
                ////Bt
                var shift_s = _context.ShiftModel.Where(d => d.code == "S").FirstOrDefault();
                var shift_c = _context.ShiftModel.Where(d => d.code == "C").FirstOrDefault();
                var utility_s = _tinhcong.GetSchedule(shift_s.id, shift_s.time_from.Value, shift_s.time_to.Value);
                var utility_c = _tinhcong.GetSchedule(shift_c.id, shift_c.time_from.Value, shift_c.time_to.Value);

                sheet.Range["C3"].DateTimeValue = date_from;
                sheet.Range["U2"].DateTimeValue = date_to;
                sheet.Range["BM10"].Value = $"Đông Hòa, ngày {date_to.ToString("dd")} tháng {date_to.ToString("MM")} năm {date_to.ToString("yy")}";


                int stt = 0;
                var start_r = 6;
                DataTable dt = new DataTable();
                dt.Columns.Add("stt", typeof(int));
                dt.Columns.Add("HOVATEN", typeof(string));
                //dt.Columns.Add("tong", typeof(decimal));
                var date_check = date_from;
                var start_column = 2;
                var end_column = 64;
                while (date_check <= date_to)
                {
                    dt.Columns.Add(date_check.ToString("yyyyMMdd") + "-S", typeof(string));
                    dt.Columns.Add(date_check.ToString("yyyyMMdd") + "-C", typeof(string));
                    date_check = date_check.AddDays(1);
                    start_column += 2;
                }
                sheet.InsertRow(8, datapost.Count(), InsertOptionsType.FormatAsAfter);

                foreach (var record in datapost)
                {
                    DataRow dr1 = dt.NewRow();
                    dr1["stt"] = (++stt);
                    dr1["HOVATEN"] = record.HOVATEN;
                    var ngaynhanviec = record.NGAYNHANVIEC;
                    var ngaynghiviec = record.NGAYNGHIVIEC;
                    //decimal tong = 0;

                    //dr1["tong"] = tong;
                    var start_c = 2;
                    date_check = date_from;

                    var nRow = sheet.Rows[start_r];

                    var phepdauky = _tinhcong.phepnam(record, date_from);

                    nRow.Cells[74].NumberValue = phepdauky.Value;
                    while (date_check <= date_to)
                    {
                        CellRange originDataRang = sheet.Range["E" + (datapost.Count + 15).ToString()]; ///NGAYNGHI
                        if (ngaynhanviec != null && ngaynhanviec > date_check)
                        {

                        }
                        else if (ngaynghiviec != null && ngaynghiviec < date_check)
                        {

                        }
                        else if (utility_s.IsAWorkDay(date_check))
                        {
                            var cong_sModel = ChamcongModel.Where(d => d.MANV == record.MANV && d.date == date_check.Date && d.shift_id == shift_s.id).FirstOrDefault();
                            if (cong_sModel == null)
                            {
                                cong_sModel = new ChamcongModel()
                                {
                                    date = date_check,
                                    MANV = record.MANV,
                                    NV_id = record.id,
                                    shift_id = shift_s.id,
                                    value = "",
                                };
                                var first_hik = list_hik.Where(d => d.id == record.MACC && d.date.Value.Date == date_check && d.time.Value >= shift_s.time_from && d.time.Value <= shift_s.time_to).FirstOrDefault();

                                if (first_hik != null)
                                {
                                    cong_sModel.value = "X";
                                }
                                if (list_holidays.Contains(date_check))
                                {
                                    cong_sModel.value = "NL";
                                }
                            }
                            switch (cong_sModel.value)
                            {
                                case "X":
                                    originDataRang = sheet.Range["E" + (datapost.Count + 11).ToString()];
                                    break;
                                case "P":
                                    originDataRang = sheet.Range["E" + (datapost.Count + 12).ToString()];
                                    break;
                                case "KL":
                                    originDataRang = sheet.Range["E" + (datapost.Count + 13).ToString()];
                                    break;
                                case "NL":
                                    originDataRang = sheet.Range["E" + (datapost.Count + 14).ToString()];
                                    break;
                                default:
                                    originDataRang = sheet.Range["E" + (datapost.Count + 16).ToString()];
                                    break;
                            }
                            dr1[date_check.ToString("yyyyMMdd") + "-S"] = cong_sModel.value;
                        }

                        var cell = nRow.Cells[start_c++];
                        sheet.Copy(originDataRang, cell, true);




                        originDataRang = sheet.Range["E" + (datapost.Count + 15).ToString()]; ///NGAYNGHI
                        if (ngaynhanviec != null && ngaynhanviec > date_check)
                        {

                        }
                        else if (ngaynghiviec != null && ngaynghiviec < date_check)
                        {

                        }
                        else if (utility_c.IsAWorkDay(date_check))
                        {
                            var cong_cModel = ChamcongModel.Where(d => d.MANV == record.MANV && d.date == date_check.Date && d.shift_id == shift_c.id).FirstOrDefault();
                            if (cong_cModel == null)
                            {
                                cong_cModel = new ChamcongModel()
                                {
                                    date = date_check,
                                    MANV = record.MANV,
                                    NV_id = record.id,
                                    shift_id = shift_c.id,
                                    value = "",
                                };
                                var first_hik = list_hik.Where(d => d.id == record.MACC && d.date.Value.Date == date_check && d.time.Value >= shift_c.time_from && d.time.Value <= shift_c.time_to).FirstOrDefault();

                                if (first_hik != null)
                                {
                                    cong_cModel.value = "X";
                                }
                                if (list_holidays.Contains(date_check))
                                {
                                    cong_cModel.value = "NL";
                                }
                            }

                            switch (cong_cModel.value)
                            {
                                case "X":
                                    originDataRang = sheet.Range["E" + (datapost.Count + 11).ToString()];
                                    break;
                                case "P":
                                    originDataRang = sheet.Range["E" + (datapost.Count + 12).ToString()];
                                    break;
                                case "KL":
                                    originDataRang = sheet.Range["E" + (datapost.Count + 13).ToString()];
                                    break;
                                case "NL":
                                    originDataRang = sheet.Range["E" + (datapost.Count + 14).ToString()];
                                    break;
                                default:
                                    originDataRang = sheet.Range["E" + (datapost.Count + 16).ToString()];
                                    break;
                            }
                            dr1[date_check.ToString("yyyyMMdd") + "-C"] = cong_cModel.value;
                        }



                        cell = nRow.Cells[start_c++];
                        sheet.Copy(originDataRang, cell, CopyRangeOptions.All);

                        date_check = date_check.AddDays(1);
                    }
                    dt.Rows.Add(dr1);
                    start_r++;
                    //CellRange originDataRang = sheet.Range["A7:BZ7"];
                    //CellRange targetDataRang = sheet.Range["A" + start_r + ":BZ" + start_r];
                    //sheet.Copy(originDataRang, targetDataRang, true);
                }

                ////// Tổng công trong tháng
                var date_working_s = utility_s.GetWorkingDaysBetweenTwoWorkingDateTimes(date_from, date_to, false);
                if (utility_s.IsAWorkDay(date_from))
                    date_working_s.Add(date_from);
                if (utility_s.IsAWorkDay(date_to))
                    date_working_s.Add(date_to);

                var date_working_c = utility_c.GetWorkingDaysBetweenTwoWorkingDateTimes(date_from, date_to, false);
                if (utility_c.IsAWorkDay(date_from))
                    date_working_c.Add(date_from);
                if (utility_c.IsAWorkDay(date_to))
                    date_working_c.Add(date_to);

                double tongcong = (date_working_s.Count() + date_working_c.Count()) / 2;

                sheet.Range["BM3"].NumberValue = tongcong;
                //sheet.Range["E145"].Value = start_r.ToString();
                sheet.InsertDataTable(dt, false, 7, 1);
                sheet.CalculateAllValue();
                var delete_count = end_column - start_column;
                if (delete_count > 0)
                {
                    sheet.DeleteColumn(start_column + 1, delete_count);
                }
            }

            ////Nhân viên DV

            datapost = datapost_all.Where(d => d.LOAIHD == "DV" && user_shift.Contains(d.id)).OrderByDescending(d => d.NGAYNHANVIEC).ToList();
            if (datapost.Count > 0)
            {
                Worksheet sheet = workbook.Worksheets[1];
                ////Bt
                var shift_s = _context.ShiftModel.Where(d => d.code == "S").FirstOrDefault();
                var shift_c = _context.ShiftModel.Where(d => d.code == "C").FirstOrDefault();
                var utility_s = _tinhcong.GetSchedule(shift_s.id, shift_s.time_from.Value, shift_s.time_to.Value);
                var utility_c = _tinhcong.GetSchedule(shift_c.id, shift_c.time_from.Value, shift_c.time_to.Value);

                sheet.Range["C3"].DateTimeValue = date_from;
                sheet.Range["U2"].DateTimeValue = date_to;
                sheet.Range["BM10"].Value = $"Đông Hòa, ngày {date_to.ToString("dd")} tháng {date_to.ToString("MM")} năm {date_to.ToString("yy")}";
                int stt = 0;
                var start_r = 6;
                DataTable dt = new DataTable();
                dt.Columns.Add("stt", typeof(int));
                dt.Columns.Add("HOVATEN", typeof(string));
                //dt.Columns.Add("tong", typeof(decimal));
                var date_check = date_from;
                var start_column = 2;
                var end_column = 64;
                while (date_check <= date_to)
                {
                    dt.Columns.Add(date_check.ToString("yyyyMMdd") + "-S", typeof(string));
                    dt.Columns.Add(date_check.ToString("yyyyMMdd") + "-C", typeof(string));
                    date_check = date_check.AddDays(1);
                    start_column += 2;
                }
                sheet.InsertRow(8, datapost.Count(), InsertOptionsType.FormatAsAfter);

                foreach (var record in datapost)
                {
                    DataRow dr1 = dt.NewRow();
                    dr1["stt"] = (++stt);
                    dr1["HOVATEN"] = record.HOVATEN;
                    var ngaynhanviec = record.NGAYNHANVIEC;
                    var ngaynghiviec = record.NGAYNGHIVIEC;
                    //decimal tong = 0;

                    //dr1["tong"] = tong;
                    var start_c = 2;
                    date_check = date_from;
                    var nRow = sheet.Rows[start_r];

                    var phepdauky = _tinhcong.phepnam(record, date_from);

                    nRow.Cells[74].NumberValue = phepdauky.Value;
                    while (date_check <= date_to)
                    {
                        CellRange originDataRang = sheet.Range["E" + (datapost.Count + 15).ToString()]; ///NGAYNGHI
                        if (ngaynhanviec != null && ngaynhanviec > date_check)
                        {

                        }
                        else if (ngaynghiviec != null && ngaynghiviec < date_check)
                        {

                        }
                        else if (utility_s.IsAWorkDay(date_check))
                        {
                            var cong_sModel = ChamcongModel.Where(d => d.MANV == record.MANV && d.date == date_check.Date && d.shift_id == shift_s.id).FirstOrDefault();
                            if (cong_sModel == null)
                            {
                                cong_sModel = new ChamcongModel()
                                {
                                    date = date_check,
                                    MANV = record.MANV,
                                    NV_id = record.id,
                                    shift_id = shift_s.id,
                                    value = "",
                                };
                                var first_hik = list_hik.Where(d => d.id == record.MACC && d.date.Value.Date == date_check && d.time.Value >= shift_s.time_from && d.time.Value <= shift_s.time_to).FirstOrDefault();

                                if (first_hik != null)
                                {
                                    cong_sModel.value = "X";
                                }
                                if (list_holidays.Contains(date_check))
                                {
                                    cong_sModel.value = "NL";
                                }
                            }
                            switch (cong_sModel.value)
                            {
                                case "X":
                                    originDataRang = sheet.Range["E" + (datapost.Count + 11).ToString()];
                                    break;
                                case "P":
                                    originDataRang = sheet.Range["E" + (datapost.Count + 12).ToString()];
                                    break;
                                case "KL":
                                    originDataRang = sheet.Range["E" + (datapost.Count + 13).ToString()];
                                    break;
                                case "NL":
                                    originDataRang = sheet.Range["E" + (datapost.Count + 14).ToString()];
                                    break;
                                default:
                                    originDataRang = sheet.Range["E" + (datapost.Count + 16).ToString()];
                                    break;
                            }
                            dr1[date_check.ToString("yyyyMMdd") + "-S"] = cong_sModel.value;
                        }

                        var cell = nRow.Cells[start_c++];
                        sheet.Copy(originDataRang, cell, true);




                        originDataRang = sheet.Range["E" + (datapost.Count + 15).ToString()]; ///NGAYNGHI
                        if (ngaynhanviec != null && ngaynhanviec > date_check)
                        {

                        }
                        else if (ngaynghiviec != null && ngaynghiviec < date_check)
                        {

                        }
                        else if (utility_c.IsAWorkDay(date_check))
                        {
                            var cong_cModel = ChamcongModel.Where(d => d.MANV == record.MANV && d.date == date_check.Date && d.shift_id == shift_c.id).FirstOrDefault();
                            if (cong_cModel == null)
                            {
                                cong_cModel = new ChamcongModel()
                                {
                                    date = date_check,
                                    MANV = record.MANV,
                                    NV_id = record.id,
                                    shift_id = shift_c.id,
                                    value = "",
                                };
                                var first_hik = list_hik.Where(d => d.id == record.MACC && d.date.Value.Date == date_check && d.time.Value >= shift_c.time_from && d.time.Value <= shift_c.time_to).FirstOrDefault();

                                if (first_hik != null)
                                {
                                    cong_cModel.value = "X";
                                }
                                if (list_holidays.Contains(date_check))
                                {
                                    cong_cModel.value = "NL";
                                }
                            }

                            switch (cong_cModel.value)
                            {
                                case "X":
                                    originDataRang = sheet.Range["E" + (datapost.Count + 11).ToString()];
                                    break;
                                case "P":
                                    originDataRang = sheet.Range["E" + (datapost.Count + 12).ToString()];
                                    break;
                                case "KL":
                                    originDataRang = sheet.Range["E" + (datapost.Count + 13).ToString()];
                                    break;
                                case "NL":
                                    originDataRang = sheet.Range["E" + (datapost.Count + 14).ToString()];
                                    break;
                                default:
                                    originDataRang = sheet.Range["E" + (datapost.Count + 16).ToString()];
                                    break;
                            }
                            dr1[date_check.ToString("yyyyMMdd") + "-C"] = cong_cModel.value;
                        }



                        cell = nRow.Cells[start_c++];
                        sheet.Copy(originDataRang, cell, CopyRangeOptions.All);

                        date_check = date_check.AddDays(1);
                    }
                    dt.Rows.Add(dr1);
                    start_r++;
                    //CellRange originDataRang = sheet.Range["A7:BZ7"];
                    //CellRange targetDataRang = sheet.Range["A" + start_r + ":BZ" + start_r];
                    //sheet.Copy(originDataRang, targetDataRang, true);
                }

                ////// Tổng công trong tháng
                var date_working_s = utility_s.GetWorkingDaysBetweenTwoWorkingDateTimes(date_from, date_to, false);
                if (utility_s.IsAWorkDay(date_from))
                    date_working_s.Add(date_from);
                if (utility_s.IsAWorkDay(date_to))
                    date_working_s.Add(date_to);

                var date_working_c = utility_c.GetWorkingDaysBetweenTwoWorkingDateTimes(date_from, date_to, false);
                if (utility_c.IsAWorkDay(date_from))
                    date_working_c.Add(date_from);
                if (utility_c.IsAWorkDay(date_to))
                    date_working_c.Add(date_to);

                double tongcong = (date_working_s.Count() + date_working_c.Count()) / 2;

                sheet.Range["BM3"].NumberValue = tongcong;
                //sheet.Range["E145"].Value = start_r.ToString();
                sheet.InsertDataTable(dt, false, 7, 1);
                sheet.CalculateAllValue();
                var delete_count = end_column - start_column;
                if (delete_count > 0)
                {
                    sheet.DeleteColumn(start_column + 1, delete_count);
                }
            }
            else
            {
                WorksheetsCollection worksheets = workbook.Worksheets;

                worksheets.Remove("Dịch vụ");

            }

            ////Nhân viên vệ sinh bảo vệ
            datapost = datapost_all.Where(d => user_shift_vs.Contains(d.id)).OrderByDescending(d => d.NGAYNHANVIEC).ToList();
            if (datapost.Count > 0)
            {
                Worksheet sheet = workbook.Worksheets["Vệ sinh-Bảo vệ-T7"];
                ////Bt
                var shift_s = _context.ShiftModel.Where(d => d.code == "S-T7").FirstOrDefault();
                var shift_c = _context.ShiftModel.Where(d => d.code == "C-T7").FirstOrDefault();
                var utility_s = _tinhcong.GetSchedule(shift_s.id, shift_s.time_from.Value, shift_s.time_to.Value);
                var utility_c = _tinhcong.GetSchedule(shift_c.id, shift_c.time_from.Value, shift_c.time_to.Value);

                sheet.Range["C3"].DateTimeValue = date_from;
                sheet.Range["U2"].DateTimeValue = date_to;
                sheet.Range["BM10"].Value = $"Đông Hòa, ngày {date_to.ToString("dd")} tháng {date_to.ToString("MM")} năm {date_to.ToString("yy")}";
                int stt = 0;
                var start_r = 6;
                DataTable dt = new DataTable();
                dt.Columns.Add("stt", typeof(int));
                dt.Columns.Add("HOVATEN", typeof(string));
                //dt.Columns.Add("tong", typeof(decimal));
                var date_check = date_from;
                var start_column = 2;
                var end_column = 64;
                while (date_check <= date_to)
                {
                    dt.Columns.Add(date_check.ToString("yyyyMMdd") + "-S", typeof(string));
                    dt.Columns.Add(date_check.ToString("yyyyMMdd") + "-C", typeof(string));
                    date_check = date_check.AddDays(1);
                    start_column += 2;
                }
                sheet.InsertRow(8, datapost.Count(), InsertOptionsType.FormatAsAfter);

                foreach (var record in datapost)
                {
                    DataRow dr1 = dt.NewRow();
                    dr1["stt"] = (++stt);
                    dr1["HOVATEN"] = record.HOVATEN;
                    var ngaynhanviec = record.NGAYNHANVIEC;
                    var ngaynghiviec = record.NGAYNGHIVIEC;
                    //decimal tong = 0;

                    //dr1["tong"] = tong;
                    var start_c = 2;
                    date_check = date_from;
                    var nRow = sheet.Rows[start_r];

                    var phepdauky = _tinhcong.phepnam(record, date_from);

                    nRow.Cells[74].NumberValue = phepdauky.Value;
                    while (date_check <= date_to)
                    {
                        CellRange originDataRang = sheet.Range["E" + (datapost.Count + 15).ToString()]; ///NGAYNGHI
                        if (ngaynhanviec != null && ngaynhanviec > date_check)
                        {

                        }
                        else if (ngaynghiviec != null && ngaynghiviec < date_check)
                        {

                        }
                        else if (utility_s.IsAWorkDay(date_check))
                        {
                            var cong_sModel = ChamcongModel.Where(d => d.MANV == record.MANV && d.date == date_check.Date && d.shift_id == shift_s.id).FirstOrDefault();
                            if (cong_sModel == null)
                            {
                                cong_sModel = new ChamcongModel()
                                {
                                    date = date_check,
                                    MANV = record.MANV,
                                    NV_id = record.id,
                                    shift_id = shift_s.id,
                                    value = "",
                                };
                                var first_hik = list_hik.Where(d => d.id == record.MACC && d.date.Value.Date == date_check && d.time.Value >= shift_s.time_from && d.time.Value <= shift_s.time_to).FirstOrDefault();

                                if (first_hik != null)
                                {
                                    cong_sModel.value = "X";
                                }
                                if (list_holidays.Contains(date_check))
                                {
                                    cong_sModel.value = "NL";
                                }
                            }
                            switch (cong_sModel.value)
                            {
                                case "X":
                                    originDataRang = sheet.Range["E" + (datapost.Count + 11).ToString()];
                                    break;
                                case "P":
                                    originDataRang = sheet.Range["E" + (datapost.Count + 12).ToString()];
                                    break;
                                case "KL":
                                    originDataRang = sheet.Range["E" + (datapost.Count + 13).ToString()];
                                    break;
                                case "NL":
                                    originDataRang = sheet.Range["E" + (datapost.Count + 14).ToString()];
                                    break;
                                default:
                                    originDataRang = sheet.Range["E" + (datapost.Count + 16).ToString()];
                                    break;
                            }
                            dr1[date_check.ToString("yyyyMMdd") + "-S"] = cong_sModel.value;
                        }

                        var cell = nRow.Cells[start_c++];
                        sheet.Copy(originDataRang, cell, true);




                        originDataRang = sheet.Range["E" + (datapost.Count + 15).ToString()]; ///NGAYNGHI
                        if (ngaynhanviec != null && ngaynhanviec > date_check)
                        {

                        }
                        else if (ngaynghiviec != null && ngaynghiviec < date_check)
                        {

                        }
                        else if (utility_c.IsAWorkDay(date_check))
                        {
                            var cong_cModel = ChamcongModel.Where(d => d.MANV == record.MANV && d.date == date_check.Date && d.shift_id == shift_c.id).FirstOrDefault();
                            if (cong_cModel == null)
                            {
                                cong_cModel = new ChamcongModel()
                                {
                                    date = date_check,
                                    MANV = record.MANV,
                                    NV_id = record.id,
                                    shift_id = shift_c.id,
                                    value = "",
                                };
                                var first_hik = list_hik.Where(d => d.id == record.MACC && d.date.Value.Date == date_check && d.time.Value >= shift_c.time_from && d.time.Value <= shift_c.time_to).FirstOrDefault();

                                if (first_hik != null)
                                {
                                    cong_cModel.value = "X";
                                }
                                if (list_holidays.Contains(date_check))
                                {
                                    cong_cModel.value = "NL";
                                }
                            }

                            switch (cong_cModel.value)
                            {
                                case "X":
                                    originDataRang = sheet.Range["E" + (datapost.Count + 11).ToString()];
                                    break;
                                case "P":
                                    originDataRang = sheet.Range["E" + (datapost.Count + 12).ToString()];
                                    break;
                                case "KL":
                                    originDataRang = sheet.Range["E" + (datapost.Count + 13).ToString()];
                                    break;
                                case "NL":
                                    originDataRang = sheet.Range["E" + (datapost.Count + 14).ToString()];
                                    break;
                                default:
                                    originDataRang = sheet.Range["E" + (datapost.Count + 16).ToString()];
                                    break;
                            }
                            dr1[date_check.ToString("yyyyMMdd") + "-C"] = cong_cModel.value;
                        }



                        cell = nRow.Cells[start_c++];
                        sheet.Copy(originDataRang, cell, CopyRangeOptions.All);

                        date_check = date_check.AddDays(1);
                    }
                    dt.Rows.Add(dr1);
                    start_r++;
                    //CellRange originDataRang = sheet.Range["A7:BZ7"];
                    //CellRange targetDataRang = sheet.Range["A" + start_r + ":BZ" + start_r];
                    //sheet.Copy(originDataRang, targetDataRang, true);
                }

                ////// Tổng công trong tháng
                var date_working_s = utility_s.GetWorkingDaysBetweenTwoWorkingDateTimes(date_from, date_to, false);
                if (utility_s.IsAWorkDay(date_from))
                    date_working_s.Add(date_from);
                if (utility_s.IsAWorkDay(date_to))
                    date_working_s.Add(date_to);

                var date_working_c = utility_c.GetWorkingDaysBetweenTwoWorkingDateTimes(date_from, date_to, false);
                if (utility_c.IsAWorkDay(date_from))
                    date_working_c.Add(date_from);
                if (utility_c.IsAWorkDay(date_to))
                    date_working_c.Add(date_to);

                double tongcong = (date_working_s.Count() + date_working_c.Count()) / 2;

                sheet.Range["BM3"].NumberValue = tongcong;
                //sheet.Range["E145"].Value = start_r.ToString();
                sheet.InsertDataTable(dt, false, 7, 1);
                sheet.CalculateAllValue(); 
                var delete_count = end_column - start_column;
                if (delete_count > 0)
                {
                    sheet.DeleteColumn(start_column + 1, delete_count);
                }
            }
            else
            {
                WorksheetsCollection worksheets = workbook.Worksheets;

                worksheets.Remove("Vệ sinh-Bảo vệ-T7");


            }
            ////Nhân viên Full 
            datapost = datapost_all.Where(d => user_shift_full.Contains(d.id)).OrderByDescending(d => d.NGAYNHANVIEC).ToList();
            if (datapost.Count > 0)
            {
                Worksheet sheet = workbook.Worksheets["Full công"];
                ////Bt
                var shift_s = _context.ShiftModel.Where(d => d.code == "F").FirstOrDefault();
                var shift_c = _context.ShiftModel.Where(d => d.code == "F").FirstOrDefault();
                var utility_s = _tinhcong.GetSchedule(shift_s.id, shift_s.time_from.Value, shift_s.time_to.Value);
                var utility_c = _tinhcong.GetSchedule(shift_c.id, shift_c.time_from.Value, shift_c.time_to.Value);

                sheet.Range["C3"].DateTimeValue = date_from;
                sheet.Range["U2"].DateTimeValue = date_to;
                sheet.Range["BM10"].Value = $"Đông Hòa, ngày {date_to.ToString("dd")} tháng {date_to.ToString("MM")} năm {date_to.ToString("yy")}";
                int stt = 0;
                var start_r = 6;
                DataTable dt = new DataTable();
                dt.Columns.Add("stt", typeof(int));
                dt.Columns.Add("HOVATEN", typeof(string));
                //dt.Columns.Add("tong", typeof(decimal));
                var date_check = date_from;
                var start_column = 2;
                var end_column = 64;
                while (date_check <= date_to)
                {
                    dt.Columns.Add(date_check.ToString("yyyyMMdd") + "-S", typeof(string));
                    dt.Columns.Add(date_check.ToString("yyyyMMdd") + "-C", typeof(string));
                    date_check = date_check.AddDays(1);
                    start_column += 2;
                }
                sheet.InsertRow(8, datapost.Count(), InsertOptionsType.FormatAsAfter);

                foreach (var record in datapost)
                {
                    DataRow dr1 = dt.NewRow();
                    dr1["stt"] = (++stt);
                    dr1["HOVATEN"] = record.HOVATEN;
                    var ngaynhanviec = record.NGAYNHANVIEC;
                    var ngaynghiviec = record.NGAYNGHIVIEC;
                    //decimal tong = 0;

                    //dr1["tong"] = tong;
                    var start_c = 2;
                    date_check = date_from;
                    var nRow = sheet.Rows[start_r];

                    var phepdauky = _tinhcong.phepnam(record, date_from);

                    nRow.Cells[74].NumberValue = phepdauky.Value;
                    while (date_check <= date_to)
                    {
                        CellRange originDataRang = sheet.Range["E" + (datapost.Count + 15).ToString()]; ///NGAYNGHI
                        if (ngaynhanviec != null && ngaynhanviec > date_check)
                        {

                        }
                        else if (ngaynghiviec != null && ngaynghiviec < date_check)
                        {

                        }
                        else if (utility_s.IsAWorkDay(date_check))
                        {
                            var cong_sModel = ChamcongModel.Where(d => d.MANV == record.MANV && d.date == date_check.Date && d.shift_id == shift_s.id).FirstOrDefault();
                            if (cong_sModel == null)
                            {
                                cong_sModel = new ChamcongModel()
                                {
                                    date = date_check,
                                    MANV = record.MANV,
                                    NV_id = record.id,
                                    shift_id = shift_s.id,
                                    value = "",
                                };
                                var first_hik = list_hik.Where(d => d.id == record.MACC && d.date.Value.Date == date_check && d.time.Value >= shift_s.time_from && d.time.Value <= shift_s.time_to).FirstOrDefault();

                                if (first_hik != null)
                                {
                                    cong_sModel.value = "X";
                                }
                                if (list_holidays.Contains(date_check))
                                {
                                    cong_sModel.value = "NL";
                                }
                            }
                            switch (cong_sModel.value)
                            {
                                case "X":
                                    originDataRang = sheet.Range["E" + (datapost.Count + 11).ToString()];
                                    break;
                                case "P":
                                    originDataRang = sheet.Range["E" + (datapost.Count + 12).ToString()];
                                    break;
                                case "KL":
                                    originDataRang = sheet.Range["E" + (datapost.Count + 13).ToString()];
                                    break;
                                case "NL":
                                    originDataRang = sheet.Range["E" + (datapost.Count + 14).ToString()];
                                    break;
                                default:
                                    originDataRang = sheet.Range["E" + (datapost.Count + 16).ToString()];
                                    break;
                            }
                            dr1[date_check.ToString("yyyyMMdd") + "-S"] = cong_sModel.value;
                        }

                        var cell = nRow.Cells[start_c++];
                        sheet.Copy(originDataRang, cell, true);




                        originDataRang = sheet.Range["E" + (datapost.Count + 15).ToString()]; ///NGAYNGHI
                        if (ngaynhanviec != null && ngaynhanviec > date_check)
                        {

                        }
                        else if (ngaynghiviec != null && ngaynghiviec < date_check)
                        {

                        }
                        else if (utility_c.IsAWorkDay(date_check))
                        {
                            var cong_cModel = ChamcongModel.Where(d => d.MANV == record.MANV && d.date == date_check.Date && d.shift_id == shift_c.id).FirstOrDefault();
                            if (cong_cModel == null)
                            {
                                cong_cModel = new ChamcongModel()
                                {
                                    date = date_check,
                                    MANV = record.MANV,
                                    NV_id = record.id,
                                    shift_id = shift_c.id,
                                    value = "",
                                };
                                var first_hik = list_hik.Where(d => d.id == record.MACC && d.date.Value.Date == date_check && d.time.Value >= shift_c.time_from && d.time.Value <= shift_c.time_to).FirstOrDefault();

                                if (first_hik != null)
                                {
                                    cong_cModel.value = "X";
                                }
                                if (list_holidays.Contains(date_check))
                                {
                                    cong_cModel.value = "NL";
                                }
                            }

                            switch (cong_cModel.value)
                            {
                                case "X":
                                    originDataRang = sheet.Range["E" + (datapost.Count + 11).ToString()];
                                    break;
                                case "P":
                                    originDataRang = sheet.Range["E" + (datapost.Count + 12).ToString()];
                                    break;
                                case "KL":
                                    originDataRang = sheet.Range["E" + (datapost.Count + 13).ToString()];
                                    break;
                                case "NL":
                                    originDataRang = sheet.Range["E" + (datapost.Count + 14).ToString()];
                                    break;
                                default:
                                    originDataRang = sheet.Range["E" + (datapost.Count + 16).ToString()];
                                    break;
                            }
                            dr1[date_check.ToString("yyyyMMdd") + "-C"] = cong_cModel.value;
                        }



                        cell = nRow.Cells[start_c++];
                        sheet.Copy(originDataRang, cell, CopyRangeOptions.All);

                        date_check = date_check.AddDays(1);
                    }
                    dt.Rows.Add(dr1);
                    start_r++;
                    //CellRange originDataRang = sheet.Range["A7:BZ7"];
                    //CellRange targetDataRang = sheet.Range["A" + start_r + ":BZ" + start_r];
                    //sheet.Copy(originDataRang, targetDataRang, true);
                }

                ////// Tổng công trong tháng
                var date_working_s = utility_s.GetWorkingDaysBetweenTwoWorkingDateTimes(date_from, date_to, false);
                if (utility_s.IsAWorkDay(date_from))
                    date_working_s.Add(date_from);
                if (utility_s.IsAWorkDay(date_to))
                    date_working_s.Add(date_to);

                var date_working_c = utility_c.GetWorkingDaysBetweenTwoWorkingDateTimes(date_from, date_to, false);
                if (utility_c.IsAWorkDay(date_from))
                    date_working_c.Add(date_from);
                if (utility_c.IsAWorkDay(date_to))
                    date_working_c.Add(date_to);

                double tongcong = (date_working_s.Count() + date_working_c.Count()) / 2;

                sheet.Range["BM3"].NumberValue = tongcong;
                //sheet.Range["E145"].Value = start_r.ToString();
                sheet.InsertDataTable(dt, false, 7, 1);
                sheet.CalculateAllValue();
                var delete_count = end_column - start_column;
                if (delete_count > 0)
                {
                    sheet.DeleteColumn(start_column + 1, delete_count);
                }
            }
            else
            {
                WorksheetsCollection worksheets = workbook.Worksheets;

                worksheets.Remove("Full công");


            }

            workbook.SaveToFile("./wwwroot" + documentPath, ExcelVersion.Version2013);
            //var congthuc_ct = _QLSXcontext.Congthuc_CTModel.Where()
            var jsonData = new { success = true, link = documentPath };
            return Json(jsonData);

            //}
            //catch (Exception ex)
            //{
            //    return Json(new { success = false, message = ex.Message });
            //}
        }
    }
    public class WHolidays : MultiCalculatedHoliDay
    {
        private List<ShiftHolidayModel> _holidays;
        internal WHolidays(List<ShiftHolidayModel> holidays) : base(0, 0)
        {
            _holidays = holidays;
        }


        public override Type GetHolyDayType() => typeof(WHolidays);

        /// <summary>Calculates the list of MultiCalculatedHoliDays for the given year.</summary>
        /// <param name="year">The year.</param>
        /// <returns></returns>
        public override List<DateTime> CalculateList(int year)
        {
            var baseHolidays = new List<DateTime>();
            baseHolidays = _holidays.Select(d => d.date.Value).ToList();
            return baseHolidays;
        }
    }
    public class ChamCong
    {
        public ShiftModel ShiftModel { get; set; }

        public DateTime Date { get; set; }
        public HikModel first_hik { get; set; }
        public ChamcongModel ChamcongModel { get; set; }

    }
}
