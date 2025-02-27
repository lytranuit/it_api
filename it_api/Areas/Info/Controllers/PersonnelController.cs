


using Holdtime.Models;
using Info.Models;
using it_template.Areas.Trend.Controllers;
using iText.Commons.Actions.Contexts;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata.Internal;
using NodaTime.TimeZones.Cldr;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Reporting;
using Spire.Xls;
using System.Collections;
using System.Data;
using System.Drawing;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
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
            var Model = _context.PersonnelModel.Where(d => d.id == id).FirstOrDefault();
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
            var EMAIL = Request.Form["filters[EMAIL]"].FirstOrDefault();
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
            if (EMAIL != null && EMAIL != "")
            {
                customerData = customerData.Where(d => d.EMAIL.Contains(EMAIL));
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
        public JsonResult Orgchart()
        {
            var data = Persons();
            data = data.Where(d => d.children.Count() > 0).ToList();
            return Json(data, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }
        private List<Personnel> Persons(string MAQUANLYTRUCTIEP = null)
        {
            var list = new List<Personnel>();

            var data = _context.PersonnelModel.Where(d => d.NGAYNGHIVIEC == null && d.MAQUANLYTRUCTIEP == MAQUANLYTRUCTIEP).ToList();
            foreach (var item in data)
            {

                var person = new Personnel()
                {
                    id = item.id,
                    name = item.HOVATEN,
                    MANV = item.MANV,
                    HOVATEN = item.HOVATEN,
                    email = item.EMAIL,
                    image_url = item.image_url
                };
                person.children = Persons(item.id);
                var chucvu = _context.PositionModel.Where(d => d.MACHUCVU == item.MACHUCVU).FirstOrDefault();
                if (chucvu != null)
                {
                    person.positionName = chucvu.TENCHUCVU;
                }
                list.Add(person);
            }
            return list;
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

        [HttpPost]
        public async Task<JsonResult> fileHDLD(string id)
        {
            var viewPath = "wwwroot/report/word/HĐLĐ mẫu.docx";
            var documentPath = "/private/info/data/" + DateTime.Now.ToFileTimeUtc() + ".docx";

            string Domain = (HttpContext.Request.IsHttps ? "https://" : "http://") + HttpContext.Request.Host.Value;

            var Model = _context.PersonnelModel.Where(d => d.id == id).FirstOrDefault();

            var chuyenmon = _context.ChuyenmonModel.Where(d => d.MACHUYENMON == Model.CHUYENMON).FirstOrDefault();
            var chucvu = _context.PositionModel.Where(d => d.MACHUCVU == Model.MACHUCVU).FirstOrDefault();
            var bophan = _context.DepartmentModel.Where(d => d.MAPHONG == Model.MAPHONG).FirstOrDefault();


            var raw = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "hovaten",Model.HOVATEN.ToUpper() },
                { "ngaysinh",Model.NGAYSINH != null ? Model.NGAYSINH.Value.ToString("dd/MM/yyyy") :" " },
                { "diachi",Model.THUONGTRU },
                { "SOCCCD",Model.SOCCCD },
                { "NGAYCAPCCCD",Model.NGAYCAPCCCD != null ? Model.NGAYCAPCCCD.Value.ToString("dd/MM/yyyy") :" " },
                { "NOICAPCCCD",Model.NOICAPCCCD },
                { "CHUYENMON",chuyenmon != null ? chuyenmon.TENCHUYENMON : " " },
                { "NGAYKYHD",Model.NGAYKYHD != null ? Model.NGAYKYHD.Value.ToString("dd/MM/yyyy") : " " },
                { "NGAYKTHD",Model.NGAYKTHD != null ? Model.NGAYKTHD.Value.ToString("dd/MM/yyyy") : " " },
                { "CHUCVU",chucvu != null ? chucvu.TENCHUCVU : " " },
                { "BOPHAN",bophan != null ? bophan.TENPHONG : " " },
                { "tien_luong",Model.tien_luong  != null ? Model.tien_luong.Value.ToString("0,###") : " " },
                { "pc_trachnhiem",Model.pc_trachnhiem != null ? Model.pc_trachnhiem.Value.ToString("0,###") : " "},

            };
            //Creates Document instance
            Spire.Doc.Document document = new Spire.Doc.Document();
            Section section = document.AddSection();

            document.LoadFromFile(viewPath, Spire.Doc.FileFormat.Docx);

            string[] fieldName = raw.Keys.ToArray();
            string[] fieldValue = raw.Values.ToArray();

            string[] MergeFieldNames = document.MailMerge.GetMergeFieldNames();
            string[] GroupNames = document.MailMerge.GetMergeGroupNames();

            document.MailMerge.Execute(fieldName, fieldValue);

            document.SaveToFile(_configuration["Source:Path_Private"] + documentPath.Replace("/private", "").Replace("/", "\\"), Spire.Doc.FileFormat.Docx);

            //var congthuc_ct = _QLSXcontext.Congthuc_CTModel.Where()
            var jsonData = new { success = true, link = documentPath };
            return Json(jsonData);
        }
        [HttpPost]
        public async Task<JsonResult> fileBMTT(string id)
        {
            var viewPath = "wwwroot/report/word/3. CK bảo mật thông tin.docx";
            var documentPath = "/private/info/data/" + DateTime.Now.ToFileTimeUtc() + ".docx";

            string Domain = (HttpContext.Request.IsHttps ? "https://" : "http://") + HttpContext.Request.Host.Value;

            var Model = _context.PersonnelModel.Where(d => d.id == id).FirstOrDefault();

            var chucvu = _context.PositionModel.Where(d => d.MACHUCVU == Model.MACHUCVU).FirstOrDefault();
            var bophan = _context.DepartmentModel.Where(d => d.MAPHONG == Model.MAPHONG).FirstOrDefault();


            var raw = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "hovaten",Model.HOVATEN.ToUpper() },
                { "ngaysinh",Model.NGAYSINH != null ? Model.NGAYSINH.Value.ToString("dd/MM/yyyy") :" " },
                { "diachi",Model.THUONGTRU },
                { "SOCCCD",Model.SOCCCD },
                { "NGAYCAPCCCD",Model.NGAYCAPCCCD != null ? Model.NGAYCAPCCCD.Value.ToString("dd/MM/yyyy") :" " },
                { "NOICAPCCCD",Model.NOICAPCCCD },
                { "CHUCVU",chucvu != null ? chucvu.TENCHUCVU : " " },
                { "BOPHAN",bophan != null ? bophan.TENPHONG : " " },
                { "NGAYNHANVIEC",Model.NGAYNHANVIEC  != null ? Model.NGAYNHANVIEC.Value.ToString("dd/MM/yyyy") : " " },

            };
            //Creates Document instance
            Spire.Doc.Document document = new Spire.Doc.Document();
            Section section = document.AddSection();

            document.LoadFromFile(viewPath, Spire.Doc.FileFormat.Docx);

            string[] fieldName = raw.Keys.ToArray();
            string[] fieldValue = raw.Values.ToArray();

            string[] MergeFieldNames = document.MailMerge.GetMergeFieldNames();
            string[] GroupNames = document.MailMerge.GetMergeGroupNames();

            document.MailMerge.Execute(fieldName, fieldValue);

            document.SaveToFile(_configuration["Source:Path_Private"] + documentPath.Replace("/private", "").Replace("/", "\\"), Spire.Doc.FileFormat.Docx);

            //var congthuc_ct = _QLSXcontext.Congthuc_CTModel.Where()
            var jsonData = new { success = true, link = documentPath };
            return Json(jsonData);
        }
        [HttpPost]
        public async Task<JsonResult> fileHVTV(string id)
        {
            var viewPath = "wwwroot/report/word/3. TT Học việc và thử việc.docx";
            var documentPath = "/private/info/data/" + DateTime.Now.ToFileTimeUtc() + ".docx";

            string Domain = (HttpContext.Request.IsHttps ? "https://" : "http://") + HttpContext.Request.Host.Value;

            var Model = _context.PersonnelModel.Where(d => d.id == id).FirstOrDefault();

            var chuyenmon = _context.ChuyenmonModel.Where(d => d.MACHUYENMON == Model.CHUYENMON).FirstOrDefault();


            var raw = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "hovaten",Model.HOVATEN.ToUpper() },
                { "ngaysinh",Model.NGAYSINH != null ? Model.NGAYSINH.Value.ToString("dd/MM/yyyy") :" " },
                { "diachi",Model.THUONGTRU },
                { "SOCCCD",Model.SOCCCD },
                { "NGAYCAPCCCD",Model.NGAYCAPCCCD != null ? Model.NGAYCAPCCCD.Value.ToString("dd/MM/yyyy") :" " },
                { "NOICAPCCCD",Model.NOICAPCCCD },
                { "CHUYENMON",chuyenmon != null ? chuyenmon.TENCHUYENMON : " " },
                { "NGAYNHANVIEC",Model.NGAYNHANVIEC != null ? Model.NGAYNHANVIEC.Value.ToString("dd/MM/yyyy") : " " }
            };
            //Creates Document instance
            Spire.Doc.Document document = new Spire.Doc.Document();
            Section section = document.AddSection();

            document.LoadFromFile(viewPath, Spire.Doc.FileFormat.Docx);

            string[] fieldName = raw.Keys.ToArray();
            string[] fieldValue = raw.Values.ToArray();

            string[] MergeFieldNames = document.MailMerge.GetMergeFieldNames();
            string[] GroupNames = document.MailMerge.GetMergeGroupNames();



            List<DictionaryEntry> relationsList = new List<DictionaryEntry>();
            if (Model.NGAYHOCVIEC != null)
            {
                relationsList.Add(new DictionaryEntry("hocviec", string.Empty));
            }
            if (Model.NGAYTHUVIEC != null)
            {
                relationsList.Add(new DictionaryEntry("thuviec", string.Empty));
            }
            //relationsList.Add(new DictionaryEntry("fres", string.Empty));
            //relationsList.Add(new DictionaryEntry("targets", "parent = %fres.key%"));
            //relationsList.Add(new DictionaryEntry("locations", "parent = %locations.key%"));

            MailMergeDataSet mailMergeDataSet = new MailMergeDataSet();
            if (Model.NGAYHOCVIEC != null)
            {
                var hocviec = new List<dynamic>()
                {
                    new {
                        NGAYHOCVIEC = Model.NGAYHOCVIEC != null ? Model.NGAYHOCVIEC.Value.ToString("dd/MM/yyyy") : " " ,
                        NGAYKTHOCVIEC = Model.NGAYKTHOCVIEC != null ? Model.NGAYKTHOCVIEC.Value.ToString("dd/MM/yyyy") : " "
                    }
                };
                mailMergeDataSet.Add(new MailMergeDataTable("hocviec", hocviec));
            }
            else
            {
                Body body = document.Sections[0].Body;
                bool insideGroup = false;
                List<Paragraph> paragraphsToRemove = new List<Paragraph>();

                foreach (Paragraph para in body.Paragraphs)
                {
                    if (para.Text.Contains("GroupStart:hocviec"))
                    {
                        insideGroup = true;
                    }

                    if (insideGroup)
                    {
                        paragraphsToRemove.Add(para);
                    }

                    if (para.Text.Contains("GroupEnd:hocviec"))
                    {
                        break;
                    }
                }

                // Xóa các đoạn văn bản trong nhóm
                foreach (Paragraph para in paragraphsToRemove)
                {
                    body.Paragraphs.Remove(para);
                }
            }
            if (Model.NGAYTHUVIEC != null)
            {
                var thuviec = new List<dynamic>()
                {
                    new {
                        NGAYTHUVIEC = Model.NGAYTHUVIEC != null ? Model.NGAYTHUVIEC.Value.ToString("dd/MM/yyyy") : " " ,
                        NGAYKTTHUVIEC = Model.NGAYKTTHUVIEC != null ? Model.NGAYKTTHUVIEC.Value.ToString("dd/MM/yyyy") : " "
                    }
                };
                mailMergeDataSet.Add(new MailMergeDataTable("thuviec", thuviec));
            }
            else
            {
                Body body = document.Sections[0].Body;
                bool insideGroup = false;
                List<Paragraph> paragraphsToRemove = new List<Paragraph>();

                foreach (Paragraph para in body.Paragraphs)
                {
                    if (para.Text.Contains("GroupStart:thuviec"))
                    {
                        insideGroup = true;
                    }

                    if (insideGroup)
                    {
                        paragraphsToRemove.Add(para);
                    }

                    if (para.Text.Contains("GroupEnd:thuviec"))
                    {
                        break;
                    }
                }

                // Xóa các đoạn văn bản trong nhóm
                foreach (Paragraph para in paragraphsToRemove)
                {
                    body.Paragraphs.Remove(para);
                }
            }


            document.MailMerge.ExecuteWidthNestedRegion(mailMergeDataSet, relationsList);


            document.MailMerge.Execute(fieldName, fieldValue);

            document.SaveToFile(_configuration["Source:Path_Private"] + documentPath.Replace("/private", "").Replace("/", "\\"), Spire.Doc.FileFormat.Docx);

            //var congthuc_ct = _QLSXcontext.Congthuc_CTModel.Where()
            var jsonData = new { success = true, link = documentPath };
            return Json(jsonData);
        }
        [HttpPost]
        public async Task<JsonResult> excel()
        {
            var viewPath = "wwwroot/report/excel/Danhsachnhansu.xlsx";
            var documentPath = "/private/info/data/" + DateTime.Now.ToFileTimeUtc() + ".xlsx";

            string Domain = (HttpContext.Request.IsHttps ? "https://" : "http://") + HttpContext.Request.Host.Value;

            Workbook workbook = new Workbook();
            workbook.LoadFromFile(viewPath);

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


            var customerData = _context.PersonnelModel.Where(d => 1 == 1);
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

            var datapost = customerData.ToList();

            var DepartmentModel = _context.DepartmentModel.ToList();
            var PositionModel = _context.PositionModel.ToList();
            var TrinhdoModel = _context.TrinhdoModel.ToList();
            var ChuyenmonModel = _context.ChuyenmonModel.ToList();
            var TinhModel = _context.TinhModel.ToList();
            datapost = datapost.Select(d =>
            {
                d.bophan = DepartmentModel.Where(e => e.MAPHONG == d.MAPHONG).FirstOrDefault();
                d.chucvu = PositionModel.Where(e => e.MACHUCVU == d.MACHUCVU).FirstOrDefault();
                d.TRINHDO = TrinhdoModel.Where(e => e.MATRINHDO == d.MATRINHDO).FirstOrDefault();
                d.CHUYENMON_Model = ChuyenmonModel.Where(e => e.MACHUYENMON == d.CHUYENMON).FirstOrDefault();
                d.TinhModel = TinhModel.Where(e => e.MaTinh == d.DIADIEM).FirstOrDefault();
                return d;
            }).OrderBy(d => d.bophan.sort)
            .ThenBy(d => d.bophan.MAPHONG)
            .ThenBy(d => d.chucvu.sort)
            .ThenBy(d => d.chucvu.MACHUCVU)
            .ThenByDescending(d => d.NGAYNHANVIEC)
            .ToList();

            Worksheet sheet = workbook.Worksheets[0];
            int start_r = 1; // Dòng gốc để sao chép định dạng

            sheet.InsertRow(start_r + 1, datapost.Count(), InsertOptionsType.FormatAsAfter);
            foreach (var record in datapost)
            {
                var nRow = sheet.Rows[start_r];
                nRow.Cells[0].Value = record.MANV.Trim();
                nRow.Cells[1].Value = record.HOVATEN.Trim();
                nRow.Cells[2].Value = record.bophan.TENPHONG.Trim();
                nRow.Cells[3].Value = record.chucvu.TENCHUCVU.Trim();
                nRow.Cells[4].Value = record.GIOITINH;
                nRow.Cells[5].Value = record.TRINHDO != null ? record.TRINHDO.TENTRINHDO.Trim() : "";
                nRow.Cells[6].Value = record.CHUYENMON_Model != null ? record.CHUYENMON_Model.TENCHUYENMON.Trim() : "";

                nRow.Cells[7].Value = record.DIENTHOAI;
                nRow.Cells[8].Value = record.EMAIL.Trim();
                if (record.NGAYSINH != null)
                {
                    nRow.Cells[9].DateTimeValue = record.NGAYSINH.Value;

                }
                if (record.NGAYNHANVIEC != null)
                {
                    nRow.Cells[10].DateTimeValue = record.NGAYNHANVIEC.Value;

                }
                if (record.NGAYNGHIVIEC != null)
                {
                    nRow.Cells[11].DateTimeValue = record.NGAYNGHIVIEC.Value;

                }
                nRow.Cells[12].Value = record.MACC;
                nRow.Cells[13].Value = record.TinhModel != null ? record.TinhModel.TenTinh.Trim() : "";
                nRow.Cells[14].Value = record.tinhtrang;
                nRow.Cells[15].Value = (record.sotk_icb ?? "") + " - " + (record.sotk_vba ?? "");
                start_r++;
            }

            workbook.SaveToFile(_configuration["Source:Path_Private"] + documentPath.Replace("/private", "").Replace("/", "\\"), ExcelVersion.Version2013);

            //var congthuc_ct = _QLSXcontext.Congthuc_CTModel.Where()
            var jsonData = new { success = true, link = documentPath };
            return Json(jsonData);
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
    public class Personnel
    {
        public string id { get; set; }
        public string type { get; set; }
        public string name { get; set; }
        public string MANV { get; set; }

        public string HOVATEN { get; set; }

        public string positionName { get; set; }
        public string image_url { get; set; }

        public string email { get; set; }
        public List<Personnel> children { get; set; }
    }
}