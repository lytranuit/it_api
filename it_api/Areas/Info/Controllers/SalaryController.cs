﻿



using Holdtime.Models;
using Info.Models;
using it_template.Areas.Trend.Controllers;
using iText.Commons.Actions.Contexts;
using iText.Commons.Utils;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using PH.WorkingDaysAndTimeUtility.Configuration;
using PH.WorkingDaysAndTimeUtility;
using System.Data;
using System.Dynamic;
using System.Text.Json.Serialization;
using Vue.Data;
using Vue.Models;
using Vue.Services;
using Spire.Xls;
using System.Security.Policy;
using Microsoft.EntityFrameworkCore.Infrastructure;
using System.Diagnostics;
using iText.StyledXmlParser.Jsoup.Nodes;

namespace it_template.Areas.Info.Controllers
{

    public class SalaryController : BaseController
    {
        private readonly IConfiguration _configuration;
        private UserManager<UserModel> UserManager;
        private readonly TinhCong _tinhcong;
        private readonly ViewRender _view;
        private readonly ItContext _itContext;

        public SalaryController(NhansuContext context, AesOperation aes, TinhCong tinhcong, IConfiguration configuration, UserManager<UserModel> UserMgr, ViewRender view, ItContext itContext) : base(context, aes)
        {
            _configuration = configuration;
            UserManager = UserMgr;
            _tinhcong = tinhcong;
            _view = view;
            _itContext = itContext;
            var listener = _itContext.GetService<DiagnosticSource>();
            (listener as DiagnosticListener).SubscribeWithAdapter(new CommandInterceptor());
        }

        [HttpPost]
        [Authorize(Roles = "Administrator,HR Lương")]
        public async Task<JsonResult> Delete(string id)
        {
            var Model = _context.SalaryModel.Where(d => d.id == id).FirstOrDefault();
            Model.deleted_at = DateTime.Now;
            _context.Update(Model);
            ///Trả lại phép năm
            if (Model.is_capnhat_phepnam == true)
            {
                var SalaryUserModel = _context.SalaryUserModel.Where(d => d.salary_id == id).ToList();
                foreach (var user in SalaryUserModel)
                {
                    var record = _context.PersonnelModel.Where(d => d.MANV == user.MANV).FirstOrDefault();
                    record.ngayphep = (double)user.phepdauky;
                    _context.Update(record);
                }
                Model.is_capnhat_phepnam = false;
            }
            _context.SaveChanges();
            return Json(new { success = true, data = Model });
        }

        [HttpPost]
        [Authorize(Roles = "Administrator,HR Lương")]
        public async Task<JsonResult> DeleteUser(int id)
        {
            var Model = _context.SalaryUserModel.Where(d => d.id == id).FirstOrDefault();

            _context.Remove(Model);
            _context.SaveChanges();
            return Json(new { success = true, data = Model });
        }
        [HttpPost]
        [Authorize(Roles = "Administrator,HR Lương")]
        public async Task<JsonResult> addUser(string salary_id, List<string> list_person)
        {
            var SalaryModel = _context.SalaryModel.Where(d => d.id == salary_id).FirstOrDefault();
            var SalaryUserModel = _context.SalaryUserModel.Where(d => d.salary_id == salary_id).Select(d => d.person_id).ToList();

            IEnumerable<string> list_add = list_person.Except(SalaryUserModel);
            if (list_add != null)
            {
                foreach (string key in list_add)
                {
                    var PersonModel = _context.PersonnelModel.Where(d => d.id == key).FirstOrDefault();
                    if (PersonModel.NGAYNHANVIEC != null && PersonModel.NGAYNHANVIEC.Value > SalaryModel.date_to.Value)
                    {
                        continue;
                    }
                    _context.Add(new SalaryUserModel()
                    {
                        salary_id = salary_id,
                        person_id = key,
                        email = PersonModel.EMAIL,

                    });
                }
                _context.SaveChanges();
            }
            return Json(new { success = true });
        }

        [HttpPost]
        [Authorize(Roles = "Administrator,HR Lương")]
        public async Task<JsonResult> tinhluong(string id)
        {
            var viewPath = "wwwroot/report/excel/Bảng lương.xlsx";
            var documentPath = "/private/info/data/" + DateTime.Now.ToFileTimeUtc() + ".xlsx";

            Workbook workbook = new Workbook();
            workbook.LoadFromFile(viewPath);



            var SalaryUserModel = _context.SalaryUserModel.Where(d => d.salary_id == id).ToList();
            var Users = SalaryUserModel.Select(d => d.person_id).ToList();
            var Model = _context.SalaryModel.Where(d => d.id == id).FirstOrDefault();
            var date_from = Model.date_from.Value;
            var date_to = Model.date_to.Value;
            var data_post = _context.PersonnelModel.Where(d => Users.Contains(d.id)).ToList();

            var list_data_cong_all = _tinhcong.cong(data_post, date_from, date_to);

            if (list_data_cong_all.Count() > 0)
            {
                //var list_data_cong = list_data_cong_all.Where(d => d.ContainsKey("MANV") && list_chinhthuc.Contains(d["MANV"])).ToList();

                var list_data_cong = list_data_cong_all.OrderBy(d => d["bophan"]).ToList();
                Worksheet sheet = workbook.Worksheets[0];
                var now = DateTime.Now;
                sheet.Range["P6"].Value = $"Tháng {date_to.ToString("MM")} Năm {date_to.ToString("yyyy")} (công tính từ ngày {date_from.ToString("dd/MM/yy")}-{date_to.ToString("dd/MM/yy")})";
                sheet.Range["AB3"].Value = $"Đông Hòa, ngày {now.ToString("dd")} tháng {now.ToString("MM")} năm {now.ToString("yyyy")}";
                int stt = 0;
                var start_r = 12;
                sheet.InsertRow(start_r + 1, list_data_cong.Count(), InsertOptionsType.FormatAsAfter);
                foreach (var item in list_data_cong)
                {
                    var person = SalaryUserModel.Where(d => d.person_id == item["NV_id"]).FirstOrDefault();

                    var record = data_post.Where(d => d.MANV == item["MANV"]).FirstOrDefault();
                    var chucvu = _context.PositionModel.Where(d => d.MACHUCVU == record.MACHUCVU).FirstOrDefault();
                    var bophan = _context.DepartmentModel.Where(d => d.MAPHONG == record.MAPHONG).FirstOrDefault();

                    var congtrongthang = item["tongcong"];
                    var congthucte = item["tong"];
                    var tenchucvu = chucvu != null ? chucvu.TENCHUCVU : "";
                    var phanloai = bophan != null ? bophan.TENPHONG : "";
                    var manv = record.MANV;
                    var hovaten = record.HOVATEN;
                    var email = record.EMAIL;
                    var luongcb = record.tien_luong ?? 0;

                    var pc_xangxe = record.pc_khac ?? 0;
                    var pc_trachnhiem = record.pc_trachnhiem ?? 0;
                    var pc_thuebang = record.pc_thuebang ?? 0;
                    var pc_khuvuc = record.pc_khuvuc ?? 0;
                    var pc_thamnien = record.pc_thamnien ?? 0;
                    var pc_hieusuat = record.pc_hieusuat ?? 0;
                    var pc_thuhut = record.pc_thuhut ?? 0;
                    var pc_khac = record.pc ?? 0;

                    var luongkpi = record.tien_luong_kpi ?? 0;
                    var TNCN_banthan = 11000000;
                    var TNCN_nguoiphuthuoc = record.nguoiphuthuoc ?? 0;
                    var tamungdot1 = record.tien_luong_dot1 ?? 0;
                    var khoantru = person.khoantru ?? 0;
                    var khoancong = person.khoancong ?? 0;
                    var note_khoantru = person.note_khoantru;
                    var note_khoancong = person.note_khoancong;
                    var note = person.note;
                    var is_bhxh = record.is_bhxh;
                    var is_thue = record.is_thue;
                    var tyle_bhxh = record.tyle_bhxh ?? 0;
                    var tyle_bhyt = record.tyle_bhyt ?? 0;
                    var tyle_bhtn = record.tyle_bhtn ?? 0;
                    var tyle_dpcd = record.tyle_dpcd ?? 0;
                    var tyle = tyle_bhxh + tyle_bhtn + tyle_bhyt;
                    var stk = (record.sotk_icb ?? "") + " - " + (record.sotk_vba ?? "");

                    var nRow = sheet.Rows[start_r];

                    if (record.LOAIHD == "DV")
                    {
                        //nRow.Cells[7].ClearAll(); // không đóng bảo hiểm

                        TNCN_banthan = 0; //// Không giảm trừ
                        TNCN_nguoiphuthuoc = 0; // Không giảm trừ
                        //tyle = 0;
                        //tyle_dpcd = 0;
                        nRow.Cells[29].Formula = "=ROUND(AC" + (start_r + 1) + " * 10%, 0)";

                        //nRow.Cells[7].Value2 = "";
                    }
                    //if (record.tinhtrang == "Học việc" || record.tinhtrang == "Thử việc" || record.tinhtrang == "Thử việc không phép")
                    //{
                    //    tyle = 0;
                    //    tyle_dpcd = 0;
                    //}
                    if (is_bhxh != true)
                    {
                        tyle = 0;
                        tyle_dpcd = 0;
                        //nRow.Cells[7].ClearAll(); // không đóng bảo hiểm
                    }
                    if (is_thue != true)
                    {
                        nRow.Cells[29].ClearAll(); // không đóng thuế TNCN
                    }

                    nRow.Cells[0].NumberValue = ++stt;
                    nRow.Cells[1].Value = manv;
                    nRow.Cells[2].Value = hovaten;
                    nRow.Cells[3].Value = tenchucvu;
                    nRow.Cells[4].Value = phanloai;
                    nRow.Cells[5].NumberValue = (double)congtrongthang;
                    nRow.Cells[6].NumberValue = (double)luongcb;

                    nRow.Cells[8].NumberValue = (double)pc_hieusuat;
                    nRow.Cells[9].NumberValue = (double)pc_thuebang;
                    nRow.Cells[10].NumberValue = (double)pc_xangxe;
                    nRow.Cells[11].NumberValue = (double)pc_thamnien;
                    nRow.Cells[12].NumberValue = (double)pc_thuhut;
                    nRow.Cells[13].NumberValue = (double)pc_khuvuc;
                    nRow.Cells[14].NumberValue = (double)pc_trachnhiem;
                    nRow.Cells[15].NumberValue = (double)pc_khac;

                    nRow.Cells[17].NumberValue = (double)luongkpi;
                    nRow.Cells[18].NumberValue = (double)congthucte;
                    nRow.Cells[19].NumberValue = (double)khoancong;
                    nRow.Cells[24].NumberValue = (double)TNCN_banthan;
                    nRow.Cells[25].NumberValue = (double)TNCN_nguoiphuthuoc;
                    nRow.Cells[31].NumberValue = (double)khoantru;
                    if (note_khoancong != null)
                        nRow.Cells[19].AddComment().Text = note_khoancong;
                    if (note_khoantru != null)
                        nRow.Cells[31].AddComment().Text = note_khoantru;
                    if (note != null)
                        nRow.Cells[20].AddComment().Text = note;


                    nRow.Cells[27].Formula = "=ROUND(H" + (start_r + 1) + " * " + tyle + "%, 0)"; ///BHXH
                    nRow.Cells[30].Formula = "=ROUND(H" + (start_r + 1) + " * " + tyle_dpcd + "%, 0)";

                    sheet.CalculateAllValue();
                    if (tamungdot1 > 0 && tamungdot1 < nRow.Cells[32].FormulaNumberValue)
                    {
                        nRow.Cells[33].NumberValue = (double)tamungdot1;
                        sheet.CalculateAllValue();
                    }
                    nRow.Cells[37].Value = stk;
                    if (record.MANV == "NVH150384")
                    {
                        Console.Write(nRow);
                    }

                    ////Cập nhật database


                    person.MANV = manv;
                    person.HOVATEN = hovaten;
                    person.CHUCVU = tenchucvu;
                    person.BOPHAN = phanloai;
                    person.email = email;
                    person.MABOPHAN = record.MAPHONG;
                    person.ngaycongchuan = congtrongthang;
                    person.ngaycongthucte = congthucte;

                    person.tyle_bhxh = (decimal)tyle_bhxh;
                    person.tyle_bhyt = (decimal)tyle_bhyt;
                    person.tyle_bhtn = (decimal)tyle_bhtn;
                    person.tyle_dpcd = (decimal)tyle_dpcd;

                    person.tc_xangxe = (decimal)pc_xangxe;
                    person.tc_chucvu = (decimal)pc_trachnhiem;
                    person.tc_hieusuat = (decimal)pc_hieusuat;
                    person.tc_thuebang = (decimal)pc_thuebang;
                    person.tc_thamnien = (decimal)pc_thamnien;
                    person.tc_thuhut = (decimal)pc_thuhut;
                    person.tc_khuvuc = (decimal)pc_khuvuc;
                    person.tc_khac = (decimal)pc_khac;

                    person.tong_tc = (decimal)Convert.ToSingle(nRow.Cells[16].FormulaValue);

                    person.luongcb = (decimal)luongcb;
                    person.luongdongbhxh = (decimal)Convert.ToSingle(nRow.Cells[7].FormulaValue);

                    person.luongkpi = (decimal)luongkpi;
                    person.tongthunhap = (decimal)Convert.ToSingle(nRow.Cells[20].FormulaValue);

                    person.thunhapchiuthue = (decimal)Convert.ToSingle(nRow.Cells[23].FormulaValue);


                    person.tncn_banthan = (decimal)TNCN_banthan;
                    person.tncn_songuoiphuthuoc = (int)TNCN_nguoiphuthuoc;
                    person.tncn_nguoiphuthuoc = (decimal)Convert.ToSingle(nRow.Cells[26].FormulaValue);
                    person.tncn_bhxh = (decimal)Convert.ToSingle(nRow.Cells[27].FormulaValue);
                    person.thunhaptinhthue = (decimal)Convert.ToSingle(nRow.Cells[28].FormulaValue);
                    person.thue_tncn = (decimal)Convert.ToSingle(nRow.Cells[29].FormulaValue);
                    person.dpcd = (decimal)Convert.ToSingle(nRow.Cells[30].FormulaValue);
                    person.thuclanh = (decimal)Convert.ToSingle(nRow.Cells[32].FormulaValue);
                    person.tamungdot1 = (decimal)Convert.ToSingle(nRow.Cells[33].EnvalutedValue);
                    person.conlai = (decimal)Convert.ToSingle(nRow.Cells[34].FormulaValue);
                    person.tongphep = (decimal)item["tongphep"];

                    decimal phepconlai = (decimal)_tinhcong.phepnam(record, date_to);
                    person.phepconlai = phepconlai;
                    person.phepdauky = phepconlai + person.tongphep;
                    //if (person.MANV == "DLT021192")
                    //{
                    //    Console.Write(1);
                    //}

                    _context.Update(person);
                    _context.SaveChanges();
                    start_r++;

                    //CellRange originDataRang = sheet.Range[$"A{(list_data_cong.Count() + 12)}:AZ{(list_data_cong.Count() + 12)}"];
                    //CellRange targetDataRang = sheet.Range["A" + start_r + ":AZ" + start_r];
                    //sheet.Copy(originDataRang, targetDataRang, true);

                }
                //sheet.InsertDataTable(dt, false, 13, 1);
                sheet.DeleteRow(12, 1);
                sheet.DeleteRow(start_r, 1);
                sheet.CalculateAllValue();
            }
            //Chính thức
            var list_chinhthuc = data_post.Where(d => d.tinhtrang == "Chính thức" && d.LOAIHD != "DV").Select(d => d.MANV).ToList();
            if (list_chinhthuc.Count() > 0)
            {
                var list_data_cong = list_data_cong_all.Where(d => d.ContainsKey("MANV") && list_chinhthuc.Contains(d["MANV"])).OrderBy(d => d["bophan"]).ToList();


                Worksheet sheet = workbook.Worksheets[1];
                var now = DateTime.Now;
                sheet.Range["P6"].Value = $"Tháng {date_to.ToString("MM")} Năm {date_to.ToString("yyyy")} (công tính từ ngày {date_from.ToString("dd/MM/yy")}-{date_to.ToString("dd/MM/yy")})";
                sheet.Range["AB3"].Value = $"Đông Hòa, ngày {now.ToString("dd")} tháng {now.ToString("MM")} năm {now.ToString("yyyy")}";
                int stt = 0;
                var start_r = 12;
                sheet.InsertRow(start_r + 1, list_data_cong.Count(), InsertOptionsType.FormatAsAfter);
                foreach (var item in list_data_cong)
                {
                    var person = SalaryUserModel.Where(d => d.person_id == item["NV_id"]).FirstOrDefault();

                    var record = data_post.Where(d => d.MANV == item["MANV"]).FirstOrDefault();
                    var chucvu = _context.PositionModel.Where(d => d.MACHUCVU == record.MACHUCVU).FirstOrDefault();
                    var bophan = _context.DepartmentModel.Where(d => d.MAPHONG == record.MAPHONG).FirstOrDefault();

                    var congtrongthang = item["tongcong"];
                    var congthucte = item["tong"];
                    var tenchucvu = chucvu != null ? chucvu.TENCHUCVU : "";
                    var phanloai = bophan != null ? bophan.TENPHONG : "";
                    var manv = record.MANV;
                    var hovaten = record.HOVATEN;
                    var email = record.EMAIL;
                    var luongcb = record.tien_luong ?? 0;

                    var pc_xangxe = record.pc_khac ?? 0;
                    var pc_trachnhiem = record.pc_trachnhiem ?? 0;
                    var pc_thuebang = record.pc_thuebang ?? 0;
                    var pc_khuvuc = record.pc_khuvuc ?? 0;
                    var pc_thamnien = record.pc_thamnien ?? 0;
                    var pc_hieusuat = record.pc_hieusuat ?? 0;
                    var pc_thuhut = record.pc_thuhut ?? 0;
                    var pc_khac = record.pc ?? 0;

                    var luongkpi = record.tien_luong_kpi ?? 0;
                    var TNCN_banthan = 11000000;
                    var TNCN_nguoiphuthuoc = record.nguoiphuthuoc ?? 0;
                    var tamungdot1 = record.tien_luong_dot1 ?? 0;
                    var khoantru = person.khoantru ?? 0;
                    var khoancong = person.khoancong ?? 0;
                    var note_khoantru = person.note_khoantru;
                    var note_khoancong = person.note_khoancong;
                    var note = person.note;
                    var is_bhxh = record.is_bhxh;
                    var is_thue = record.is_thue;
                    var tyle_bhxh = record.tyle_bhxh ?? 0;
                    var tyle_bhyt = record.tyle_bhyt ?? 0;
                    var tyle_bhtn = record.tyle_bhtn ?? 0;
                    var tyle_dpcd = record.tyle_dpcd ?? 0;
                    var tyle = tyle_bhxh + tyle_bhtn + tyle_bhyt;
                    var stk = (record.sotk_icb ?? "") + " - " + (record.sotk_vba ?? "");

                    var nRow = sheet.Rows[start_r];

                    if (record.LOAIHD == "DV")
                    {
                        //nRow.Cells[7].ClearAll(); // không đóng bảo hiểm

                        TNCN_banthan = 0; //// Không giảm trừ
                        TNCN_nguoiphuthuoc = 0; // Không giảm trừ
                        //tyle = 0;
                        //tyle_dpcd = 0;
                        nRow.Cells[29].Formula = "=ROUND(AC" + (start_r + 1) + " * 10%, 0)";

                        //nRow.Cells[7].Value2 = "";
                    }
                    //if (record.tinhtrang == "Học việc" || record.tinhtrang == "Thử việc" || record.tinhtrang == "Thử việc không phép")
                    //{
                    //    tyle = 0;
                    //    tyle_dpcd = 0;
                    //}
                    if (is_bhxh != true)
                    {
                        tyle = 0;
                        tyle_dpcd = 0;
                        //nRow.Cells[7].ClearAll(); // không đóng bảo hiểm
                    }
                    if (is_thue != true)
                    {
                        nRow.Cells[29].ClearAll(); // không đóng thuế TNCN
                    }

                    nRow.Cells[0].NumberValue = ++stt;
                    nRow.Cells[1].Value = manv;
                    nRow.Cells[2].Value = hovaten;
                    nRow.Cells[3].Value = tenchucvu;
                    nRow.Cells[4].Value = phanloai;
                    nRow.Cells[5].NumberValue = (double)congtrongthang;
                    nRow.Cells[6].NumberValue = (double)luongcb;

                    nRow.Cells[8].NumberValue = (double)pc_hieusuat;
                    nRow.Cells[9].NumberValue = (double)pc_thuebang;
                    nRow.Cells[10].NumberValue = (double)pc_xangxe;
                    nRow.Cells[11].NumberValue = (double)pc_thamnien;
                    nRow.Cells[12].NumberValue = (double)pc_thuhut;
                    nRow.Cells[13].NumberValue = (double)pc_khuvuc;
                    nRow.Cells[14].NumberValue = (double)pc_trachnhiem;
                    nRow.Cells[15].NumberValue = (double)pc_khac;

                    nRow.Cells[17].NumberValue = (double)luongkpi;
                    nRow.Cells[18].NumberValue = (double)congthucte;
                    nRow.Cells[19].NumberValue = (double)khoancong;
                    nRow.Cells[24].NumberValue = (double)TNCN_banthan;
                    nRow.Cells[25].NumberValue = (double)TNCN_nguoiphuthuoc;
                    nRow.Cells[31].NumberValue = (double)khoantru;
                    if (note_khoancong != null)
                        nRow.Cells[19].AddComment().Text = note_khoancong;
                    if (note_khoantru != null)
                        nRow.Cells[31].AddComment().Text = note_khoantru;
                    if (note != null)
                        nRow.Cells[20].AddComment().Text = note;


                    nRow.Cells[27].Formula = "=ROUND(H" + (start_r + 1) + " * " + tyle + "%, 0)"; ///BHXH
                    nRow.Cells[30].Formula = "=ROUND(H" + (start_r + 1) + " * " + tyle_dpcd + "%, 0)";

                    sheet.CalculateAllValue();
                    if (tamungdot1 > 0 && tamungdot1 < nRow.Cells[32].FormulaNumberValue)
                    {
                        nRow.Cells[33].NumberValue = (double)tamungdot1;
                        sheet.CalculateAllValue();
                    }
                    nRow.Cells[37].Value = stk;

                    start_r++;

                    //CellRange originDataRang = sheet.Range[$"A{(list_data_cong.Count() + 12)}:AZ{(list_data_cong.Count() + 12)}"];
                    //CellRange targetDataRang = sheet.Range["A" + start_r + ":AZ" + start_r];
                    //sheet.Copy(originDataRang, targetDataRang, true);

                }
                //sheet.InsertDataTable(dt, false, 13, 1);
                sheet.DeleteRow(12, 1);
                sheet.DeleteRow(start_r, 1);
                sheet.CalculateAllValue();
            }
            ///Dich vu
            var list_dichvu = data_post.Where(d => d.LOAIHD == "DV").Select(d => d.MANV).ToList();
            if (list_dichvu.Count() > 0)
            {
                var list_data_cong = list_data_cong_all.Where(d => d.ContainsKey("MANV") && list_dichvu.Contains(d["MANV"])).OrderBy(d => d["bophan"]).ToList();


                Worksheet sheet = workbook.Worksheets[2];
                var now = DateTime.Now;
                sheet.Range["P6"].Value = $"Tháng {date_to.ToString("MM")} Năm {date_to.ToString("yyyy")} (công tính từ ngày {date_from.ToString("dd/MM/yy")}-{date_to.ToString("dd/MM/yy")})";
                sheet.Range["AB3"].Value = $"Đông Hòa, ngày {now.ToString("dd")} tháng {now.ToString("MM")} năm {now.ToString("yyyy")}";
                int stt = 0;
                var start_r = 12;
                sheet.InsertRow(start_r + 1, list_data_cong.Count(), InsertOptionsType.FormatAsAfter);
                foreach (var item in list_data_cong)
                {
                    var person = SalaryUserModel.Where(d => d.person_id == item["NV_id"]).FirstOrDefault();

                    var record = data_post.Where(d => d.MANV == item["MANV"]).FirstOrDefault();
                    var chucvu = _context.PositionModel.Where(d => d.MACHUCVU == record.MACHUCVU).FirstOrDefault();
                    var bophan = _context.DepartmentModel.Where(d => d.MAPHONG == record.MAPHONG).FirstOrDefault();

                    var congtrongthang = item["tongcong"];
                    var congthucte = item["tong"];
                    var tenchucvu = chucvu != null ? chucvu.TENCHUCVU : "";
                    var phanloai = bophan != null ? bophan.TENPHONG : "";
                    var manv = record.MANV;
                    var hovaten = record.HOVATEN;
                    var email = record.EMAIL;
                    var luongcb = record.tien_luong ?? 0;

                    var pc_xangxe = record.pc_khac ?? 0;
                    var pc_trachnhiem = record.pc_trachnhiem ?? 0;
                    var pc_thuebang = record.pc_thuebang ?? 0;
                    var pc_khuvuc = record.pc_khuvuc ?? 0;
                    var pc_thamnien = record.pc_thamnien ?? 0;
                    var pc_hieusuat = record.pc_hieusuat ?? 0;
                    var pc_thuhut = record.pc_thuhut ?? 0;
                    var pc_khac = record.pc ?? 0;

                    var luongkpi = record.tien_luong_kpi ?? 0;
                    var TNCN_banthan = 11000000;
                    var TNCN_nguoiphuthuoc = record.nguoiphuthuoc ?? 0;
                    var tamungdot1 = record.tien_luong_dot1 ?? 0;
                    var khoantru = person.khoantru ?? 0;
                    var khoancong = person.khoancong ?? 0;
                    var note_khoantru = person.note_khoantru;
                    var note_khoancong = person.note_khoancong;
                    var note = person.note;
                    var is_bhxh = record.is_bhxh;
                    var is_thue = record.is_thue;
                    var tyle_bhxh = record.tyle_bhxh ?? 0;
                    var tyle_bhyt = record.tyle_bhyt ?? 0;
                    var tyle_bhtn = record.tyle_bhtn ?? 0;
                    var tyle_dpcd = record.tyle_dpcd ?? 0;
                    var tyle = tyle_bhxh + tyle_bhtn + tyle_bhyt;
                    var stk = (record.sotk_icb ?? "") + " - " + (record.sotk_vba ?? "");

                    var nRow = sheet.Rows[start_r];

                    if (record.LOAIHD == "DV")
                    {
                        //nRow.Cells[7].ClearAll(); // không đóng bảo hiểm

                        TNCN_banthan = 0; //// Không giảm trừ
                        TNCN_nguoiphuthuoc = 0; // Không giảm trừ
                        //tyle = 0;
                        //tyle_dpcd = 0;
                        nRow.Cells[29].Formula = "=ROUND(AC" + (start_r + 1) + " * 10%, 0)";

                        //nRow.Cells[7].Value2 = "";
                    }
                    //if (record.tinhtrang == "Học việc" || record.tinhtrang == "Thử việc" || record.tinhtrang == "Thử việc không phép")
                    //{
                    //    tyle = 0;
                    //    tyle_dpcd = 0;
                    //}
                    if (is_bhxh != true)
                    {
                        tyle = 0;
                        tyle_dpcd = 0;
                        //nRow.Cells[7].ClearAll(); // không đóng bảo hiểm
                    }
                    if (is_thue != true)
                    {
                        nRow.Cells[29].ClearAll(); // không đóng thuế TNCN
                    }

                    nRow.Cells[0].NumberValue = ++stt;
                    nRow.Cells[1].Value = manv;
                    nRow.Cells[2].Value = hovaten;
                    nRow.Cells[3].Value = tenchucvu;
                    nRow.Cells[4].Value = phanloai;
                    nRow.Cells[5].NumberValue = (double)congtrongthang;
                    nRow.Cells[6].NumberValue = (double)luongcb;

                    nRow.Cells[8].NumberValue = (double)pc_hieusuat;
                    nRow.Cells[9].NumberValue = (double)pc_thuebang;
                    nRow.Cells[10].NumberValue = (double)pc_xangxe;
                    nRow.Cells[11].NumberValue = (double)pc_thamnien;
                    nRow.Cells[12].NumberValue = (double)pc_thuhut;
                    nRow.Cells[13].NumberValue = (double)pc_khuvuc;
                    nRow.Cells[14].NumberValue = (double)pc_trachnhiem;
                    nRow.Cells[15].NumberValue = (double)pc_khac;

                    nRow.Cells[17].NumberValue = (double)luongkpi;
                    nRow.Cells[18].NumberValue = (double)congthucte;
                    nRow.Cells[19].NumberValue = (double)khoancong;
                    nRow.Cells[24].NumberValue = (double)TNCN_banthan;
                    nRow.Cells[25].NumberValue = (double)TNCN_nguoiphuthuoc;
                    nRow.Cells[31].NumberValue = (double)khoantru;
                    if (note_khoancong != null)
                        nRow.Cells[19].AddComment().Text = note_khoancong;
                    if (note_khoantru != null)
                        nRow.Cells[31].AddComment().Text = note_khoantru;
                    if (note != null)
                        nRow.Cells[20].AddComment().Text = note;


                    nRow.Cells[27].Formula = "=ROUND(H" + (start_r + 1) + " * " + tyle + "%, 0)"; ///BHXH
                    nRow.Cells[30].Formula = "=ROUND(H" + (start_r + 1) + " * " + tyle_dpcd + "%, 0)";

                    sheet.CalculateAllValue();
                    if (tamungdot1 > 0 && tamungdot1 < nRow.Cells[32].FormulaNumberValue)
                    {
                        nRow.Cells[33].NumberValue = (double)tamungdot1;
                        sheet.CalculateAllValue();
                    }
                    nRow.Cells[37].Value = stk;


                    start_r++;

                    //CellRange originDataRang = sheet.Range[$"A{(list_data_cong.Count() + 12)}:AZ{(list_data_cong.Count() + 12)}"];
                    //CellRange targetDataRang = sheet.Range["A" + start_r + ":AZ" + start_r];
                    //sheet.Copy(originDataRang, targetDataRang, true);

                }
                //sheet.InsertDataTable(dt, false, 13, 1);
                sheet.DeleteRow(12, 1);
                sheet.DeleteRow(start_r, 1);
                sheet.CalculateAllValue();
            }

            ///Học việc thử việc
            var list_hocviec_thuviec = data_post.Where(d => (d.tinhtrang == "Học việc" || d.tinhtrang == "Thử việc" || d.tinhtrang == "Thử việc không bảo hiểm (BH)") && d.LOAIHD != "DV").Select(d => d.MANV).ToList();
            if (list_hocviec_thuviec.Count() > 0)
            {
                var list_data_cong = list_data_cong_all.Where(d => d.ContainsKey("MANV") && list_hocviec_thuviec.Contains(d["MANV"])).OrderBy(d => d["bophan"]).ToList();


                Worksheet sheet = workbook.Worksheets[3];
                var now = DateTime.Now;
                sheet.Range["P6"].Value = $"Tháng {date_to.ToString("MM")} Năm {date_to.ToString("yyyy")} (công tính từ ngày {date_from.ToString("dd/MM/yy")}-{date_to.ToString("dd/MM/yy")})";
                sheet.Range["AB3"].Value = $"Đông Hòa, ngày {now.ToString("dd")} tháng {now.ToString("MM")} năm {now.ToString("yyyy")}";
                int stt = 0;
                var start_r = 12;
                sheet.InsertRow(start_r + 1, list_data_cong.Count(), InsertOptionsType.FormatAsAfter);
                foreach (var item in list_data_cong)
                {
                    var person = SalaryUserModel.Where(d => d.person_id == item["NV_id"]).FirstOrDefault();

                    var record = data_post.Where(d => d.MANV == item["MANV"]).FirstOrDefault();
                    var chucvu = _context.PositionModel.Where(d => d.MACHUCVU == record.MACHUCVU).FirstOrDefault();
                    var bophan = _context.DepartmentModel.Where(d => d.MAPHONG == record.MAPHONG).FirstOrDefault();

                    var congtrongthang = item["tongcong"];
                    var congthucte = item["tong"];
                    var tenchucvu = chucvu != null ? chucvu.TENCHUCVU : "";
                    var phanloai = bophan != null ? bophan.TENPHONG : "";
                    var manv = record.MANV;
                    var hovaten = record.HOVATEN;
                    var email = record.EMAIL;
                    var luongcb = record.tien_luong ?? 0;

                    var pc_xangxe = record.pc_khac ?? 0;
                    var pc_trachnhiem = record.pc_trachnhiem ?? 0;
                    var pc_thuebang = record.pc_thuebang ?? 0;
                    var pc_khuvuc = record.pc_khuvuc ?? 0;
                    var pc_thamnien = record.pc_thamnien ?? 0;
                    var pc_hieusuat = record.pc_hieusuat ?? 0;
                    var pc_thuhut = record.pc_thuhut ?? 0;
                    var pc_khac = record.pc ?? 0;

                    var luongkpi = record.tien_luong_kpi ?? 0;
                    var TNCN_banthan = 11000000;
                    var TNCN_nguoiphuthuoc = record.nguoiphuthuoc ?? 0;
                    var tamungdot1 = record.tien_luong_dot1 ?? 0;
                    var khoantru = person.khoantru ?? 0;
                    var khoancong = person.khoancong ?? 0;
                    var note_khoantru = person.note_khoantru;
                    var note_khoancong = person.note_khoancong;
                    var note = person.note;
                    var is_bhxh = record.is_bhxh;
                    var is_thue = record.is_thue;
                    var tyle_bhxh = record.tyle_bhxh ?? 0;
                    var tyle_bhyt = record.tyle_bhyt ?? 0;
                    var tyle_bhtn = record.tyle_bhtn ?? 0;
                    var tyle_dpcd = record.tyle_dpcd ?? 0;
                    var tyle = tyle_bhxh + tyle_bhtn + tyle_bhyt;
                    var stk = (record.sotk_icb ?? "") + " - " + (record.sotk_vba ?? "");

                    var nRow = sheet.Rows[start_r];

                    if (record.LOAIHD == "DV")
                    {
                        //nRow.Cells[7].ClearAll(); // không đóng bảo hiểm

                        TNCN_banthan = 0; //// Không giảm trừ
                        TNCN_nguoiphuthuoc = 0; // Không giảm trừ
                        //tyle = 0;
                        //tyle_dpcd = 0;
                        nRow.Cells[29].Formula = "=ROUND(AC" + (start_r + 1) + " * 10%, 0)";

                        //nRow.Cells[7].Value2 = "";
                    }
                    //if (record.tinhtrang == "Học việc" || record.tinhtrang == "Thử việc" || record.tinhtrang == "Thử việc không phép")
                    //{
                    //    tyle = 0;
                    //    tyle_dpcd = 0;
                    //}
                    if (is_bhxh != true)
                    {
                        tyle = 0;
                        tyle_dpcd = 0;
                        //nRow.Cells[7].ClearAll(); // không đóng bảo hiểm
                    }
                    if (is_thue != true)
                    {
                        nRow.Cells[29].ClearAll(); // không đóng thuế TNCN
                    }

                    nRow.Cells[0].NumberValue = ++stt;
                    nRow.Cells[1].Value = manv;
                    nRow.Cells[2].Value = hovaten;
                    nRow.Cells[3].Value = tenchucvu;
                    nRow.Cells[4].Value = phanloai;
                    nRow.Cells[5].NumberValue = (double)congtrongthang;
                    nRow.Cells[6].NumberValue = (double)luongcb;

                    nRow.Cells[8].NumberValue = (double)pc_hieusuat;
                    nRow.Cells[9].NumberValue = (double)pc_thuebang;
                    nRow.Cells[10].NumberValue = (double)pc_xangxe;
                    nRow.Cells[11].NumberValue = (double)pc_thamnien;
                    nRow.Cells[12].NumberValue = (double)pc_thuhut;
                    nRow.Cells[13].NumberValue = (double)pc_khuvuc;
                    nRow.Cells[14].NumberValue = (double)pc_trachnhiem;
                    nRow.Cells[15].NumberValue = (double)pc_khac;

                    nRow.Cells[17].NumberValue = (double)luongkpi;
                    nRow.Cells[18].NumberValue = (double)congthucte;
                    nRow.Cells[19].NumberValue = (double)khoancong;
                    nRow.Cells[24].NumberValue = (double)TNCN_banthan;
                    nRow.Cells[25].NumberValue = (double)TNCN_nguoiphuthuoc;
                    nRow.Cells[31].NumberValue = (double)khoantru;
                    if (note_khoancong != null)
                        nRow.Cells[19].AddComment().Text = note_khoancong;
                    if (note_khoantru != null)
                        nRow.Cells[31].AddComment().Text = note_khoantru;
                    if (note != null)
                        nRow.Cells[20].AddComment().Text = note;


                    nRow.Cells[27].Formula = "=ROUND(H" + (start_r + 1) + " * " + tyle + "%, 0)"; ///BHXH
                    nRow.Cells[30].Formula = "=ROUND(H" + (start_r + 1) + " * " + tyle_dpcd + "%, 0)";

                    sheet.CalculateAllValue();
                    if (tamungdot1 > 0 && tamungdot1 < nRow.Cells[32].FormulaNumberValue)
                    {
                        nRow.Cells[33].NumberValue = (double)tamungdot1;
                        sheet.CalculateAllValue();
                    }
                    nRow.Cells[37].Value = stk;

                    start_r++;

                    //CellRange originDataRang = sheet.Range[$"A{(list_data_cong.Count() + 12)}:AZ{(list_data_cong.Count() + 12)}"];
                    //CellRange targetDataRang = sheet.Range["A" + start_r + ":AZ" + start_r];
                    //sheet.Copy(originDataRang, targetDataRang, true);

                }
                //sheet.InsertDataTable(dt, false, 13, 1);
                sheet.DeleteRow(12, 1);
                sheet.DeleteRow(start_r, 1);
                sheet.CalculateAllValue();
            }

            //Từng bộ phận
            var list_bophan = list_data_cong_all.GroupBy(d => d["bophan"]).Select(d => new
            {
                key = d.Key,
                list = d.ToList(),
            }).ToList();
            foreach (var item1 in list_bophan)
            {
                string bp = item1.key;
                var list_data_cong = item1.list;
                var bophan = _context.DepartmentModel.Where(d => d.MAPHONG == bp).FirstOrDefault();

                Worksheet Osheet = workbook.Worksheets[4];
                Worksheet sheet = workbook.CreateEmptySheet(bophan.TENPHONG);
                sheet.CopyFrom(Osheet);

                var now = DateTime.Now;
                sheet.Range["P6"].Value = $"Tháng {date_to.ToString("MM")} Năm {date_to.ToString("yyyy")} (công tính từ ngày {date_from.ToString("dd/MM/yy")}-{date_to.ToString("dd/MM/yy")})";
                sheet.Range["AB3"].Value = $"Đông Hòa, ngày {now.ToString("dd")} tháng {now.ToString("MM")} năm {now.ToString("yyyy")}";
                int stt = 0;
                var start_r = 12;
                sheet.InsertRow(start_r + 1, list_data_cong.Count(), InsertOptionsType.FormatAsAfter);
                foreach (var item in list_data_cong)
                {
                    var person = SalaryUserModel.Where(d => d.person_id == item["NV_id"]).FirstOrDefault();

                    var record = data_post.Where(d => d.MANV == item["MANV"]).FirstOrDefault();
                    var chucvu = _context.PositionModel.Where(d => d.MACHUCVU == record.MACHUCVU).FirstOrDefault();
                    //var bophan = _context.DepartmentModel.Where(d => d.MAPHONG == record.MAPHONG).FirstOrDefault();

                    var congtrongthang = item["tongcong"];
                    var congthucte = item["tong"];
                    var tenchucvu = chucvu != null ? chucvu.TENCHUCVU : "";
                    var phanloai = bophan != null ? bophan.TENPHONG : "";
                    var manv = record.MANV;
                    var hovaten = record.HOVATEN;
                    var email = record.EMAIL;
                    var luongcb = record.tien_luong ?? 0;

                    var pc_xangxe = record.pc_khac ?? 0;
                    var pc_trachnhiem = record.pc_trachnhiem ?? 0;
                    var pc_thuebang = record.pc_thuebang ?? 0;
                    var pc_khuvuc = record.pc_khuvuc ?? 0;
                    var pc_thamnien = record.pc_thamnien ?? 0;
                    var pc_hieusuat = record.pc_hieusuat ?? 0;
                    var pc_thuhut = record.pc_thuhut ?? 0;
                    var pc_khac = record.pc ?? 0;

                    var luongkpi = record.tien_luong_kpi ?? 0;
                    var TNCN_banthan = 11000000;
                    var TNCN_nguoiphuthuoc = record.nguoiphuthuoc ?? 0;
                    var tamungdot1 = record.tien_luong_dot1 ?? 0;
                    var khoantru = person.khoantru ?? 0;
                    var khoancong = person.khoancong ?? 0;
                    var note_khoantru = person.note_khoantru;
                    var note_khoancong = person.note_khoancong;
                    var note = person.note;
                    var is_bhxh = record.is_bhxh;
                    var is_thue = record.is_thue;
                    var tyle_bhxh = record.tyle_bhxh ?? 0;
                    var tyle_bhyt = record.tyle_bhyt ?? 0;
                    var tyle_bhtn = record.tyle_bhtn ?? 0;
                    var tyle_dpcd = record.tyle_dpcd ?? 0;
                    var tyle = tyle_bhxh + tyle_bhtn + tyle_bhyt;
                    var stk = (record.sotk_icb ?? "") + " - " + (record.sotk_vba ?? "");

                    var nRow = sheet.Rows[start_r];

                    if (record.LOAIHD == "DV")
                    {
                        //nRow.Cells[7].ClearAll(); // không đóng bảo hiểm

                        TNCN_banthan = 0; //// Không giảm trừ
                        TNCN_nguoiphuthuoc = 0; // Không giảm trừ
                        //tyle = 0;
                        //tyle_dpcd = 0;
                        nRow.Cells[29].Formula = "=ROUND(AC" + (start_r + 1) + " * 10%, 0)";

                        //nRow.Cells[7].Value2 = "";
                    }
                    //if (record.tinhtrang == "Học việc" || record.tinhtrang == "Thử việc" || record.tinhtrang == "Thử việc không phép")
                    //{
                    //    tyle = 0;
                    //    tyle_dpcd = 0;
                    //}
                    if (is_bhxh != true)
                    {
                        tyle = 0;
                        tyle_dpcd = 0;
                        //nRow.Cells[7].ClearAll(); // không đóng bảo hiểm
                    }
                    if (is_thue != true)
                    {
                        nRow.Cells[29].ClearAll(); // không đóng thuế TNCN
                    }

                    nRow.Cells[0].NumberValue = ++stt;
                    nRow.Cells[1].Value = manv;
                    nRow.Cells[2].Value = hovaten;
                    nRow.Cells[3].Value = tenchucvu;
                    nRow.Cells[4].Value = phanloai;
                    nRow.Cells[5].NumberValue = (double)congtrongthang;
                    nRow.Cells[6].NumberValue = (double)luongcb;

                    nRow.Cells[8].NumberValue = (double)pc_hieusuat;
                    nRow.Cells[9].NumberValue = (double)pc_thuebang;
                    nRow.Cells[10].NumberValue = (double)pc_xangxe;
                    nRow.Cells[11].NumberValue = (double)pc_thamnien;
                    nRow.Cells[12].NumberValue = (double)pc_thuhut;
                    nRow.Cells[13].NumberValue = (double)pc_khuvuc;
                    nRow.Cells[14].NumberValue = (double)pc_trachnhiem;
                    nRow.Cells[15].NumberValue = (double)pc_khac;

                    nRow.Cells[17].NumberValue = (double)luongkpi;
                    nRow.Cells[18].NumberValue = (double)congthucte;
                    nRow.Cells[19].NumberValue = (double)khoancong;
                    nRow.Cells[24].NumberValue = (double)TNCN_banthan;
                    nRow.Cells[25].NumberValue = (double)TNCN_nguoiphuthuoc;
                    nRow.Cells[31].NumberValue = (double)khoantru;
                    if (note_khoancong != null)
                        nRow.Cells[19].AddComment().Text = note_khoancong;
                    if (note_khoantru != null)
                        nRow.Cells[31].AddComment().Text = note_khoantru;
                    if (note != null)
                        nRow.Cells[20].AddComment().Text = note;


                    nRow.Cells[27].Formula = "=ROUND(H" + (start_r + 1) + " * " + tyle + "%, 0)"; ///BHXH
                    nRow.Cells[30].Formula = "=ROUND(H" + (start_r + 1) + " * " + tyle_dpcd + "%, 0)";

                    sheet.CalculateAllValue();
                    if (tamungdot1 > 0 && tamungdot1 < nRow.Cells[32].FormulaNumberValue)
                    {
                        nRow.Cells[33].NumberValue = (double)tamungdot1;
                        sheet.CalculateAllValue();
                    }
                    nRow.Cells[37].Value = stk;

                    start_r++;

                    //CellRange originDataRang = sheet.Range[$"A{(list_data_cong.Count() + 12)}:AZ{(list_data_cong.Count() + 12)}"];
                    //CellRange targetDataRang = sheet.Range["A" + start_r + ":AZ" + start_r];
                    //sheet.Copy(originDataRang, targetDataRang, true);

                }
                //sheet.InsertDataTable(dt, false, 13, 1);
                sheet.DeleteRow(12, 1);
                sheet.DeleteRow(start_r, 1);
                sheet.CalculateAllValue();
            }
            workbook.Worksheets.Remove("Clone");

            workbook.SaveToFile(_configuration["Source:Path_Private"] + documentPath.Replace("/private", "").Replace("/", "\\"), ExcelVersion.Version2013);

            Model.file_url = documentPath;
            _context.Update(Model);
            _context.SaveChanges();
            //var congthuc_ct = _QLSXcontext.Congthuc_CTModel.Where()
            var jsonData = new { success = true };
            return Json(jsonData);

        }
        [HttpPost]
        [Authorize(Roles = "Administrator,HR Lương")]
        public async Task<JsonResult> capnhatphepnam(string id)
        {
            var SalaryModel = _context.SalaryModel.Where(d => d.id == id).FirstOrDefault();
            SalaryModel.is_capnhat_phepnam = true;
            var SalaryUserModel = _context.SalaryUserModel.Where(d => d.salary_id == id).ToList();
            foreach (var user in SalaryUserModel)
            {
                var record = _context.PersonnelModel.Where(d => d.MANV == user.MANV).FirstOrDefault();
                record.ngayphep = (double)user.phepconlai;
                record.ngayphep_date = SalaryModel.date_to;
                _context.Update(record);
            }
            _context.SaveChanges();
            var jsonData = new { success = true };
            return Json(jsonData);
        }
        [HttpPost]
        [Authorize(Roles = "Administrator,HR Lương")]
        public async Task<JsonResult> tralaiphepnam(string id)
        {
            var SalaryModel = _context.SalaryModel.Where(d => d.id == id).FirstOrDefault();
            SalaryModel.is_capnhat_phepnam = false;
            var SalaryUserModel = _context.SalaryUserModel.Where(d => d.salary_id == id).ToList();
            foreach (var user in SalaryUserModel)
            {
                var record = _context.PersonnelModel.Where(d => d.MANV == user.MANV).FirstOrDefault();
                record.ngayphep = (double)user.phepdauky - 1;
                record.ngayphep_date = SalaryModel.date_from.Value.AddDays(-1);
                _context.Update(record);
            }

            _context.SaveChanges();
            var jsonData = new { success = true };
            return Json(jsonData);
        }
        [HttpPost]
        [Authorize(Roles = "Administrator,HR Lương")]
        public async Task<JsonResult> sendphieuluong(string id)
        {
            var SalaryUserModel = _context.SalaryUserModel.Where(d => d.salary_id == id).ToList();
            var Model = _context.SalaryModel.Where(d => d.id == id).FirstOrDefault();
            foreach (var s in SalaryUserModel)
            {
                var documentPath = _tinhcong.phieuluong(s.id, _configuration);
                var person = _context.PersonnelModel.Where(d => d.MANV == s.MANV).FirstOrDefault();

                string Domain = (HttpContext.Request.IsHttps ? "https://" : "http://") + HttpContext.Request.Host.Value;
                var body = _view.Render("Emails/Phieuluong", new
                {
                    link_logo = Domain + "/images/clientlogo_astahealthcare.com_f1800.png",
                    link = _configuration["Application:Info:link"] + "salary",
                    date_to = Model.date_to.Value
                });
                var attach = new List<string>()
                {
                    documentPath
                };
                var email = new EmailModel
                {
                    email_to = person.EMAIL,
                    subject = "[Phiếu lương] " + Model.name,
                    body = body,
                    email_type = "Phieuluong",
                    status = 1,
                    data_attachments = attach
                };
                _itContext.Add(email);
            }
            _itContext.SaveChanges();

            var jsonData = new { success = true };
            return Json(jsonData);

        }
        [HttpPost]
        public async Task<JsonResult> phieuluong(int id)
        {
            var documentPath = _tinhcong.phieuluong(id, _configuration);

            //var congthuc_ct = _QLSXcontext.Congthuc_CTModel.Where()
            var jsonData = new { success = true, link = documentPath };
            return Json(jsonData);

        }
        [HttpPost]
        [Authorize(Roles = "Administrator,HR Lương")]
        public async Task<JsonResult> Save(SalaryModel SalaryModel, List<string> list_person)
        {
            System.Security.Claims.ClaimsPrincipal currentUser = this.User;
            var user_id = UserManager.GetUserId(currentUser);
            var user = await UserManager.GetUserAsync(currentUser);
            SalaryModel? SalaryModel_old;
            var properties = typeof(SalaryModel).GetProperties().Where(prop => prop.CanRead && prop.CanWrite && prop.PropertyType == typeof(DateTime?));

            foreach (var prop in properties)
            {
                DateTime? value = (DateTime?)prop.GetValue(SalaryModel, null);
                if (value != null && value.Value.Kind == DateTimeKind.Utc)
                {
                    value = value.Value.ToLocalTime();
                    prop.SetValue(SalaryModel, value, null);
                }
            }
            if (SalaryModel.date != null)
            {
                SalaryModel.nam = SalaryModel.date.Value.Year;
                SalaryModel.thang = SalaryModel.date.Value.Month;
            }


            if (SalaryModel.id == null)
            {
                SalaryModel.id = Guid.NewGuid().ToString();
                SalaryModel.created_at = DateTime.Now;

                SalaryModel.created_by = user_id;


                _context.SalaryModel.Add(SalaryModel);

                _context.SaveChanges();

                SalaryModel_old = SalaryModel;

                //////
                var list_person_old = _context.SalaryUserModel.Where(d => d.salary_id == SalaryModel_old.id).Select(d => d.person_id).ToList();
                IEnumerable<string> list_delete_user = list_person_old.Except(list_person);
                IEnumerable<string> list_add_user = list_person.Except(list_person_old);

                if (list_add_user != null)
                {
                    foreach (string key in list_add_user)
                    {
                        var PersonModel = _context.PersonnelModel.Where(d => d.id == key).FirstOrDefault();
                        if (PersonModel.NGAYNHANVIEC != null && PersonModel.NGAYNHANVIEC.Value > SalaryModel.date_to.Value)
                        {
                            continue;
                        }
                        _context.Add(new SalaryUserModel()
                        {
                            salary_id = SalaryModel_old.id,
                            person_id = key,
                            email = PersonModel.EMAIL,

                        });
                    }
                    //_context.SaveChanges();
                }
                if (list_delete_user != null)
                {
                    foreach (string key in list_delete_user)
                    {
                        SalaryUserModel SalaryUserModel = _context.SalaryUserModel.Where(d => d.salary_id == SalaryModel_old.id && d.person_id == key).First();
                        _context.Remove(SalaryUserModel);
                    }
                    //_context.SaveChanges();
                }
                _context.SaveChanges();
            }
            else
            {

                SalaryModel_old = _context.SalaryModel.Where(d => d.id == SalaryModel.id).FirstOrDefault();
                CopyValues<SalaryModel>(SalaryModel_old, SalaryModel);
                SalaryModel_old.updated_at = DateTime.Now;

                _context.Update(SalaryModel_old);
                _context.SaveChanges();
            }

            return Json(new { success = true, data = SalaryModel_old });
        }

        [HttpPost]
        [Authorize(Roles = "Administrator,HR Lương")]
        public async Task<JsonResult> SaveSalaryUser(SalaryUserModel SalaryUserModel)
        {
            var SalaryUserModel_old = _context.SalaryUserModel.Where(d => d.id == SalaryUserModel.id).FirstOrDefault();
            if (ModelState.ContainsKey("khoantru"))
            {
                SalaryUserModel_old.khoantru = SalaryUserModel.khoantru;
            }
            if (ModelState.ContainsKey("khoancong"))
            {
                SalaryUserModel_old.khoancong = SalaryUserModel.khoancong;
            }
            if (ModelState.ContainsKey("note_khoancong"))
            {
                SalaryUserModel_old.note_khoancong = SalaryUserModel.note_khoancong;
            }
            if (ModelState.ContainsKey("note_khoantru"))
            {
                SalaryUserModel_old.note_khoantru = SalaryUserModel.note_khoantru;
            }
            if (ModelState.ContainsKey("note"))
            {
                SalaryUserModel_old.note = SalaryUserModel.note;
            }
            _context.Update(SalaryUserModel_old);
            _context.SaveChanges();


            return Json(new { success = true });
        }

        [HttpPost]
        [Authorize(Roles = "Administrator,HR Lương")]
        public async Task<JsonResult> ApplyAll(SalaryUserModel SalaryUserModel)
        {
            var list_user = _context.SalaryUserModel.Where(d => d.salary_id == SalaryUserModel.salary_id).ToList();

            foreach (var item in list_user)
            {
                if (ModelState.ContainsKey("khoantru"))
                {

                    item.khoantru = SalaryUserModel.khoantru;
                }
                if (ModelState.ContainsKey("khoancong"))
                {
                    item.khoancong = SalaryUserModel.khoancong;
                }
                if (ModelState.ContainsKey("note_khoancong"))
                {
                    item.note_khoancong = SalaryUserModel.note_khoancong;
                }
                if (ModelState.ContainsKey("note_khoantru"))
                {
                    item.note_khoantru = SalaryUserModel.note_khoantru;
                }
            }

            _context.UpdateRange(list_user);
            _context.SaveChanges();


            return Json(new { success = true });
        }




        [HttpPost]
        public async Task<JsonResult> Table()
        {
            var draw = Request.Form["draw"].FirstOrDefault();
            var start = Request.Form["start"].FirstOrDefault();
            var length = Request.Form["length"].FirstOrDefault();
            int pageSize = length != null ? Convert.ToInt32(length) : 0;
            var description = Request.Form["filters[description]"].FirstOrDefault();
            var name = Request.Form["filters[name]"].FirstOrDefault();
            var id = Request.Form["filters[id]"].FirstOrDefault();
            //var tenhh = Request.Form["filters[tenhh]"].FirstOrDefault();
            int skip = start != null ? Convert.ToInt32(start) : 0;
            var customerData = _context.SalaryModel.Where(d => d.deleted_at == null);
            int recordsTotal = customerData.Count();
            if (description != null && description != "")
            {
                customerData = customerData.Where(d => d.description.Contains(description));
            }
            if (name != null && name != "")
            {
                customerData = customerData.Where(d => d.name.Contains(name));
            }
            if (id != null)
            {
                customerData = customerData.Where(d => d.id == id);
            }

            int recordsFiltered = customerData.Count();
            var datapost = customerData.OrderByDescending(d => d.created_at).Skip(skip).Take(pageSize).ToList();
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

        public JsonResult Get(string id)
        {
            var data = _context.SalaryModel.Where(d => d.id == id).FirstOrDefault();
            return Json(data, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }
        public async Task<JsonResult> GetUser(string id)
        {
            var data = _context.SalaryUserModel.Where(d => d.salary_id == id).ToList();
            return Json(data, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }
        public async Task<JsonResult> listMySalary()
        {

            System.Security.Claims.ClaimsPrincipal currentUser = this.User;
            var user_id = UserManager.GetUserId(currentUser);
            var user = await UserManager.GetUserAsync(currentUser);
            var person = _context.PersonnelModel.Where(d => d.EMAIL.ToLower() == user.Email.ToLower()).FirstOrDefault();
            var MANV = person.MANV;

            var data = _context.SalaryUserModel.Where(d => d.MANV == MANV).Include(d => d.salary).Select(d => d.salary).ToList();
            data = data.Where(d => d.deleted_at == null && d.status == "Đã khóa").ToList();
            return Json(data, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }
        public async Task<JsonResult> GetMySalary(string id)
        {
            System.Security.Claims.ClaimsPrincipal currentUser = this.User;
            var user_id = UserManager.GetUserId(currentUser);
            var user = await UserManager.GetUserAsync(currentUser);
            var person = _context.PersonnelModel.Where(d => d.EMAIL == user.Email).FirstOrDefault();
            var MANV = person.MANV;


            var data = _context.SalaryUserModel.Where(d => d.salary_id == id && d.MANV == MANV).FirstOrDefault();
            return Json(data, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }

    }

}
