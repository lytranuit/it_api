


using Holdtime.Models;
using Info.Models;
using it_template.Areas.Trend.Controllers;
using iText.Commons.Actions.Contexts;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using NodaTime.TimeZones.Cldr;
using Spire.Xls;
using System.Data;
using System.Drawing;
using System.Text.Json.Serialization;
using Vue.Data;
using Vue.Models;
using Vue.Services;

namespace it_template.Areas.Info.Controllers
{

    [Authorize(Roles = "Administrator,HR")]
    public class PersonnelController : BaseController
    {
        private readonly IConfiguration _configuration;
        private UserManager<UserModel> UserManager;
        public PersonnelController(NhansuContext context, AesOperation aes, IConfiguration configuration, UserManager<UserModel> UserMgr) : base(context, aes)
        {
            _configuration = configuration;
            UserManager = UserMgr;
        }

        [HttpPost]
        public async Task<JsonResult> Delete(string id)
        {
            var Model = _context.PersonnelModel.Where(d => d.MANV == id).FirstOrDefault();
            _context.Remove(Model);
            _context.SaveChanges();
            return Json(new { success = true, data = Model });
        }
        [HttpPost]
        public async Task<JsonResult> Save(PersonnelModel PersonnelModel, List<string> list_shift)
        {
            System.Security.Claims.ClaimsPrincipal currentUser = this.User;
            var user_id = UserManager.GetUserId(currentUser);
            var user = await UserManager.GetUserAsync(currentUser);
            PersonnelModel? PersonnelModel_old;
            var properties = typeof(PersonnelModel).GetProperties().Where(prop => prop.CanRead && prop.CanWrite && prop.PropertyType == typeof(DateTime?));

            foreach (var prop in properties)
            {
                DateTime? value = (DateTime?)prop.GetValue(PersonnelModel, null);
                if (value != null && value.Value.Kind == DateTimeKind.Utc)
                {
                    value = value.Value.ToLocalTime();
                    prop.SetValue(PersonnelModel, value, null);
                }
            }
            if (PersonnelModel.id == null)
            {
                PersonnelModel.id = Guid.NewGuid().ToString();


                _context.PersonnelModel.Add(PersonnelModel);

                _context.SaveChanges();

                PersonnelModel_old = PersonnelModel;

            }
            else
            {

                PersonnelModel_old = _context.PersonnelModel.Where(d => d.id == PersonnelModel.id).FirstOrDefault();
                CopyValues<PersonnelModel>(PersonnelModel_old, PersonnelModel);

                _context.Update(PersonnelModel_old);
                _context.SaveChanges();
            }
            ///// 
            //////
            if (list_shift != null && list_shift.Count() > 0)
            {
                var list_shift_old = _context.ShiftUserModel.Where(d => d.person_id == PersonnelModel_old.id).Select(d => d.shift_id).ToList();
                IEnumerable<string> list_delete = list_shift_old.Except(list_shift);
                IEnumerable<string> list_add = list_shift.Except(list_shift_old);

                if (list_add != null)
                {
                    foreach (string key in list_add)
                    {

                        _context.Add(new ShiftUserModel()
                        {
                            shift_id = key,
                            person_id = PersonnelModel_old.id,
                            email = PersonnelModel_old.EMAIL,

                        });
                    }
                    //_context.SaveChanges();
                }
                if (list_delete != null)
                {
                    foreach (string key in list_delete)
                    {
                        ShiftUserModel ShiftUserModel = _context.ShiftUserModel.Where(d => d.person_id == PersonnelModel_old.id && d.shift_id == key).First();
                        _context.Remove(ShiftUserModel);
                    }
                    //_context.SaveChanges();
                }
                _context.SaveChanges();
            }

            return Json(new { success = true, data = PersonnelModel_old });
        }
        [HttpPost]
        public async Task<JsonResult> SaveAutoEat(bool auto)
        {
            System.Security.Claims.ClaimsPrincipal currentUser = this.User;
            var user_id = UserManager.GetUserId(currentUser);
            var user = await UserManager.GetUserAsync(currentUser);
            var email = user.Email;

            var person = _context.PersonnelModel.Where(d => d.EMAIL == email).FirstOrDefault();
            if (person != null)
            {
                person.autoeat = auto;
                _context.Update(person);
                _context.SaveChanges();
                return Json(new { success = true });
            }
            else
            {
                return Json(new { success = false, message = "Không tìm thấy nhân viên" });
            }

        }

        [HttpPost]
        public async Task<JsonResult> Table()
        {
            var draw = Request.Form["draw"].FirstOrDefault();
            var start = Request.Form["start"].FirstOrDefault();
            var length = Request.Form["length"].FirstOrDefault();
            int pageSize = length != null ? Convert.ToInt32(length) : 0;
            var ignore_nghi = Request.Form["filters[ignore_nghi]"].FirstOrDefault();
            var MANV = Request.Form["filters[MANV]"].FirstOrDefault();
            var HOVATEN = Request.Form["filters[HOVATEN]"].FirstOrDefault();
            var GIOITINH = Request.Form["filters[GIOITINH]"].FirstOrDefault();
            var tinhtrang = Request.Form["filters[tinhtrang]"].FirstOrDefault();
            var MAPHONG = Request.Form["filters[MAPHONG]"].FirstOrDefault();
            var MATRINHDO = Request.Form["filters[MATRINHDO]"].FirstOrDefault();
            var CHUYENMON = Request.Form["filters[CHUYENMON]"].FirstOrDefault();
            var BOPHAN = Request.Form["filters[BOPHAN]"].FirstOrDefault();
            var id = Request.Form["filters[id]"].FirstOrDefault();
            //var tenhh = Request.Form["filters[tenhh]"].FirstOrDefault();
            int skip = start != null ? Convert.ToInt32(start) : 0;
            var customerData = _context.PersonnelModel.Where(d => 1 == 1);
            int recordsTotal = customerData.Count();
            if (ignore_nghi == "true")
            {
                customerData = customerData.Where(d => d.NGAYNGHIVIEC == null);
            }
            if (MANV != null && MANV != "")
            {
                customerData = customerData.Where(d => d.MANV.Contains(MANV));
            }
            if (HOVATEN != null && HOVATEN != "")
            {
                customerData = customerData.Where(d => d.HOVATEN.Contains(HOVATEN));
            }
            if (GIOITINH != null && GIOITINH != "")
            {
                customerData = customerData.Where(d => d.GIOITINH == GIOITINH);
            }

            if (tinhtrang != null && tinhtrang != "")
            {
                customerData = customerData.Where(d => d.tinhtrang == tinhtrang);
            }

            if (MAPHONG != null && MAPHONG != "")
            {
                customerData = customerData.Where(d => d.MAPHONG == MAPHONG);
            }

            if (MATRINHDO != null && MATRINHDO != "")
            {
                customerData = customerData.Where(d => d.MATRINHDO == MATRINHDO);
            }

            if (CHUYENMON != null && CHUYENMON != "")
            {
                customerData = customerData.Where(d => d.CHUYENMON == CHUYENMON);
            }
            if (BOPHAN != null && BOPHAN != "")
            {
                customerData = customerData.Where(d => d.MAPHONG == BOPHAN);
            }
            if (id != null)
            {
                customerData = customerData.Where(d => d.id == id);
            }

            int recordsFiltered = customerData.Count();
            var datapost = customerData.OrderByDescending(d => d.NGAYNHANVIEC).ThenBy(d => d.MANV).Skip(skip).Take(pageSize).ToList();
            //var data = new ArrayList();
            foreach (var record in datapost)
            {
                record.list_shift = _context.ShiftUserModel.Where(d => d.person_id == record.id).Select(d => d.shift_id).ToList();
            }


            var jsonData = new { draw = draw, recordsFiltered = recordsFiltered, recordsTotal = recordsTotal, data = datapost };
            return Json(jsonData, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }
        public async Task<JsonResult> sync()
        {
            //return Ok();
            // Khởi tạo workbook để đọc
            Spire.Xls.Workbook book = new Spire.Xls.Workbook();
            book.LoadFromFile("./wwwroot/report/excel/Personnel File.xlsx", ExcelVersion.Version2013);

            Spire.Xls.Worksheet sheet = book.Worksheets[0];
            var lastrow = sheet.LastDataRow;
            // nếu vẫn chưa gặp end thì vẫn lấy data    
            Console.WriteLine(lastrow);
            var list_Result = new List<ResultModel>();
            var location_id = 1;
            var location = "";
            var stt = 0;
            var list_person = _context.PersonnelModel.ToList();
            for (int rowIndex = 1; rowIndex < lastrow; rowIndex++)
            {
                // lấy row hiện tại
                var nowRow = sheet.Rows[rowIndex];
                if (nowRow == null)
                    continue;
                // vì ta dùng 3 cột A, B, C => data của ta sẽ như sau
                //int numcount = nowRow.Cells.Count;
                //for(int y = 0;y<numcount - 1 ;y++)
                var macc = nowRow.Cells[0] != null ? nowRow.Cells[0].Value.TrimStart('\'') : null;
                // Xuất ra thông tin lên màn hình
                Console.WriteLine("MS: {0} ", macc);
                Console.WriteLine("rowIndex: {0} ", rowIndex);

                if (macc == null)
                    continue;

                var name = nowRow.Cells[2] != null ? nowRow.Cells[2].Value : null;
                name = name.ToLower().Trim();
                PersonnelModel? findP = null;
                //DateTime? date = nowRow.Cells[2] != null ? nowRow.Cells[2].DateTimeValue : null;
                foreach (var p in list_person)
                {
                    var hoten = RemoveUnicode(p.HOVATEN).ToLower().Trim();
                    if (hoten == name)
                    {
                        findP = p;
                        break;
                    }
                }

                if (findP != null)
                {
                    findP.MACC = macc;
                    _context.Update(findP);
                    _context.SaveChanges();
                }



                //EquipmentModel EquipmentModel = new EquipmentModel { code = code, name = name_vn, name_en = name_en, created_at = DateTime.Now };
                //_context.Add(EquipmentModel);
                //_context.SaveChanges();
            }
            //_context.AddRange(list_Result);
            //_context.SaveChanges();

            return Json(new { success = true });
        }
        public JsonResult Get(string id)
        {
            var data = _context.PersonnelModel.Where(d => d.id == id).FirstOrDefault();
            data.list_shift = _context.ShiftUserModel.Where(d => d.person_id == data.id).Select(d => d.shift_id).ToList();
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



        private string RemoveUnicode(string text)
        {
            string[] arr1 = new string[] { "á", "à", "ả", "ã", "ạ", "â", "ấ", "ầ", "ẩ", "ẫ", "ậ", "ă", "ắ", "ằ", "ẳ", "ẵ", "ặ",
    "đ",
    "é","è","ẻ","ẽ","ẹ","ê","ế","ề","ể","ễ","ệ",
    "í","ì","ỉ","ĩ","ị",
    "ó","ò","ỏ","õ","ọ","ô","ố","ồ","ổ","ỗ","ộ","ơ","ớ","ờ","ở","ỡ","ợ",
    "ú","ù","ủ","ũ","ụ","ư","ứ","ừ","ử","ữ","ự",
    "ý","ỳ","ỷ","ỹ","ỵ",};
            string[] arr2 = new string[] { "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a",
    "d",
    "e","e","e","e","e","e","e","e","e","e","e",
    "i","i","i","i","i",
    "o","o","o","o","o","o","o","o","o","o","o","o","o","o","o","o","o",
    "u","u","u","u","u","u","u","u","u","u","u",
    "y","y","y","y","y",};
            for (int i = 0; i < arr1.Length; i++)
            {
                text = text.Replace(arr1[i], arr2[i]);
                text = text.Replace(arr1[i].ToUpper(), arr2[i].ToUpper());
            }
            return text;
        }
    }

}
