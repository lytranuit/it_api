﻿using it_template.Areas.Info.Controllers;
using PH.WorkingDaysAndTimeUtility.Configuration;
using PH.WorkingDaysAndTimeUtility;
using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using Vue.Data;
using System.Dynamic;
using static System.Runtime.InteropServices.JavaScript.JSType;
using Info.Models;
using Microsoft.EntityFrameworkCore;
using Spire.Xls;
using WorkTimeCalculatorLib.Models;
using WorkTimeCalculatorLib;

namespace Vue.Services
{
    public class TinhCong
    {

        private readonly byte[] iv;
        private readonly string key;
        protected readonly NhansuContext _context;
        public TinhCong(NhansuContext context)
        {
            _context = context;
        }
        public string phieuluong(int id, IConfiguration _configuration)
        {
            var SalaryUserModel = _context.SalaryUserModel.Where(d => d.id == id).Include(d => d.salary).FirstOrDefault();

            var viewPath = "wwwroot/report/excel/Phieuluong.xlsx";
            var documentPath = "/private/info/data/" + DateTime.Now.ToFileTimeUtc() + ".xlsx";
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(viewPath);
            Worksheet sheet = workbook.Worksheets[0];
            var now = DateTime.Now;
            var date = SalaryUserModel.salary.date_to.Value;
            var manv = SalaryUserModel.MANV;
            var tennv = SalaryUserModel.HOVATEN;
            var chucdanh = SalaryUserModel.CHUCVU;

            var tyle_bhxh = SalaryUserModel.tyle_bhxh;
            var tyle_bhyt = SalaryUserModel.tyle_bhyt;
            var tyle_bhtn = SalaryUserModel.tyle_bhtn;
            var tyle_dpcd = SalaryUserModel.tyle_dpcd;

            sheet.Range["H1"].DateTimeValue = date;
            sheet.Range["H2"].DateTimeValue = now;
            sheet.Range["C6"].Value = manv;
            sheet.Range["C7"].Value = tennv;
            sheet.Range["C8"].Value = chucdanh;
            sheet.Range["C10"].NumberValue = (double)(SalaryUserModel.luongcb ?? 0);
            sheet.Range["C12"].NumberValue = (double)(SalaryUserModel.tc_xangxe ?? 0);
            sheet.Range["C13"].NumberValue = (double)(SalaryUserModel.tc_hieusuat ?? 0);
            sheet.Range["C14"].NumberValue = (double)(SalaryUserModel.tc_chucvu ?? 0);
            sheet.Range["C15"].NumberValue = (double)((SalaryUserModel.tc_khac ?? 0) + (SalaryUserModel.tc_thuhut ?? 0) + (SalaryUserModel.tc_khuvuc ?? 0) + (SalaryUserModel.tc_thamnien ?? 0) + (SalaryUserModel.tc_thuebang ?? 0));
            sheet.Range["C16"].NumberValue = (double)(SalaryUserModel.luongkpi ?? 0);
            sheet.Range["C17"].NumberValue = (double)(SalaryUserModel.khoancong ?? 0);
            //sheet.Range["C18"].NumberValue = (double)(SalaryUserModel.tongthunhap ?? 0);

            sheet.Range["F6"].NumberValue = (double)(SalaryUserModel.luongdongbhxh ?? 0);
            sheet.Range["F7"].NumberValue = (double)(SalaryUserModel.ngaycongthucte ?? 0);
            sheet.Range["F8"].NumberValue = (double)(SalaryUserModel.ngaycongchuan ?? 0);

            sheet.Range["F11"].Formula = "=$F$6 *" + tyle_bhxh + "%";
            sheet.Range["F12"].Formula = "=$F$6 *" + tyle_bhyt + "%";
            sheet.Range["F13"].Formula = "=$F$6 *" + tyle_bhtn + "%";

            sheet.Range["F14"].NumberValue = (double)(SalaryUserModel.thue_tncn ?? 0);

            sheet.Range["F16"].Formula = "=$F$6 *" + tyle_dpcd + "%";

            sheet.Range["D19"].NumberValue = (double)((SalaryUserModel.thuclanh ?? 0) + (SalaryUserModel.khoantru ?? 0));
            sheet.Range["D20"].NumberValue = (double)(SalaryUserModel.khoantru ?? 0);
            //sheet.Range["D20"].NumberValue = (double)(SalaryUserModel.thuclanh ?? 0);
            sheet.Range["D22"].NumberValue = (double)(SalaryUserModel.tamungdot1 ?? 0);
            //sheet.Range["D23"].NumberValue = (double)(SalaryUserModel.conlai ?? 0);

            sheet.CalculateAllValue();
            workbook.SaveToFile(_configuration["Source:Path_Private"] + documentPath.Replace("/private", "").Replace("/", "\\"), ExcelVersion.Version2013);

            SalaryUserModel.file_url = documentPath;
            _context.Update(SalaryUserModel);
            _context.SaveChanges();
            return documentPath;
        }
        public List<IDictionary<string, dynamic>> cong(List<PersonnelModel> datapost, DateTime date_from, DateTime date_to, bool addCongMoi = false)
        {
            var data = new List<IDictionary<string, dynamic>>();
            var list_nv = datapost.Select(d => d.MANV).ToList();
            var ChamcongModel = _context.ChamcongModel.Where(d => d.date.Value.Date >= date_from && d.date.Value.Date <= date_to && list_nv.Contains(d.MANV)).ToList();
            var list_hik = _context.HikModel.Where(d => d.date.Value.Date >= date_from && d.date.Value.Date <= date_to && d.device != "A.CT.1").OrderBy(d => d.date).ThenBy(d => d.time).ToList();
            var list_holidays = _context.HolidayModel.Where(d => d.date.Value.Date >= date_from && d.date.Value.Date <= date_to).Select(d => d.date).ToList();
            foreach (var record in datapost)
            {
                var d = new ExpandoObject() as IDictionary<string, dynamic>;
                var list_chamcong = new List<ChamCong>();
                var shifts = _context.ShiftUserModel.Where(d => d.person_id == record.id).Include(d => d.shift).Select(d => d.shift).ToList();
                var CongMoi = new List<ChamcongModel>();
                DateTime? ngaynhanviec = record.NGAYNHANVIEC != null ? record.NGAYNHANVIEC.Value.Date : null;
                DateTime? ngaynghiviec = record.NGAYNGHIVIEC != null ? record.NGAYNGHIVIEC.Value.Date : null;
                //if (ngaynghiviec != null)
                //{
                //    Console.WriteLine(ngaynghiviec.ToString());
                //}
                decimal tong = 0;
                decimal tongphep = 0;
                foreach (var shift in shifts)
                {

                    var utility = GetSchedule(shift.id, shift.time_from.Value, shift.time_to.Value, date_from, date_to);
                    var date_working = utility.GetWorkingDaysBetweenTwoWorkingDateTimes(date_from, date_to, false);
                    //var date_working = utility.CalculateWorkTime(date_from, date_to, false);
                    if (utility.IsAWorkDay(date_from))
                        date_working.Add(date_from);
                    if (utility.IsAWorkDay(date_to))
                        date_working.Add(date_to);

                    var cong = new ExpandoObject() as IDictionary<string, dynamic>;
                    var t = 0;
                    foreach (var date in date_working)
                    {
                        var first_hik = list_hik.Where(d => d.id == record.MACC && d.date.Value.Date == date.Date && d.time.Value >= shift.time_from && d.time.Value <= shift.time_to).FirstOrDefault();
                        list_chamcong.Add(new ChamCong()
                        {
                            ShiftModel = shift,
                            Date = date,
                            first_hik = first_hik,

                        });


                    }
                }
                var chamcong = list_chamcong.GroupBy(c => c.Date).Select(c => new
                {
                    key = c.Key.ToString("yyyyMMdd"),
                    date = c.Key,
                    shifts = c.ToList(),
                    //shift = c.GroupBy(e => e.ShiftModel).Select(s => new
                    //{
                    //    shift = s.Key,
                    //    data = s.Select(e => e.list_hiks).ToList(),
                    //})
                }).ToList();
                bool is_finish_chamcong = true;
                foreach (var c in chamcong)
                {
                    if (ngaynhanviec != null && ngaynhanviec > c.date)
                    {
                        continue;
                    }
                    if (ngaynghiviec != null && ngaynghiviec <= c.date)
                    {
                        continue;
                    }
                    var value = new ExpandoObject() as IDictionary<string, dynamic>;
                    decimal cong = 0;
                    decimal phep = 0;
                    foreach (var shift in c.shifts)
                    {
                        var congModel = ChamcongModel.Where(d => d.MANV == record.MANV && d.date == c.date && d.shift_id == shift.ShiftModel.id).FirstOrDefault();
                        if (congModel == null)
                        {
                            congModel = new ChamcongModel()
                            {
                                date = c.date,
                                MANV = record.MANV,
                                NV_id = record.id,
                                shift_id = shift.ShiftModel.id,
                                value = "",
                                value_new = "",
                                is_duyet = true,
                            };
                            if (shift.first_hik != null)
                            {
                                congModel.value = "X";
                                congModel.value_new = "X";
                            }
                            if (list_holidays.Contains(c.date))
                            {
                                congModel.value = "NL";
                                congModel.value_new = "NL";
                            }
                            if (addCongMoi)
                            {
                                CongMoi.Add(congModel);
                            }
                        }
                        if (congModel.value == "" || congModel.value != congModel.value_new)
                        {
                            is_finish_chamcong = false;
                        }
                        if (congModel.value == "X" || congModel.value == "NL" || congModel.value == "Pr" || congModel.value == "P")
                        {
                            cong += shift.ShiftModel.factor != null ? (decimal)shift.ShiftModel.factor : 1;
                        }
                        if (congModel.value == "NX")
                        {
                            cong += shift.ShiftModel.factor != null ? (decimal)(shift.ShiftModel.factor) * 0.7m : 0.7m;
                        }
                        if (congModel.value == "P")
                        {
                            phep += shift.ShiftModel.factor != null ? (decimal)shift.ShiftModel.factor : 1;
                        }
                        shift.ChamcongModel = congModel;
                    }
                    tong += cong;
                    tongphep += phep;
                    value.Add("cong", cong);
                    value.Add("phep", phep);
                    value.Add("shifts", c.shifts);
                    d.Add(c.key, value);
                }

                var phepconlai = (decimal)phepnam(record, date_to);
                DateTime denngay = new DateTime(date_from.Year, 12, 25); /// Cuối năm
                if (date_from > denngay)
                {
                    denngay = new DateTime(date_from.Year + 1, 12, 25);
                }
                var pheptoida = (decimal)phepnam(record, denngay);
                d.Add("NV_id", record.id);
                d.Add("MANV", record.MANV);
                d.Add("HOVATEN", record.HOVATEN);
                d.Add("bophan", record.MAPHONG);
                d.Add("EMAIL", record.EMAIL);
                d.Add("MACHUCVU", record.MACHUCVU);
                d.Add("tong", tong);
                d.Add("is_finish_chamcong", is_finish_chamcong);
                d.Add("tongphep", tongphep);
                d.Add("phepnamconlai", phepconlai);
                d.Add("phepnamtoida", pheptoida);
                d.Add("tongcong", chamcong.Count());
                if (addCongMoi)
                {
                    _context.AddRange(CongMoi);
                    _context.SaveChanges();
                }

                data.Add(d);
            }
            return data;
        }

        public double? phepnam(PersonnelModel person, DateTime denngay, bool include = true)
        {
            var tungay = person.ngayphep_date;
            var phep = person.ngayphep;
            double? sophepnam = null;
            if (person.tinhtrang != "Chính thức" && person.tinhtrang != "Thử việc" && person.tinhtrang != "Dịch vụ")
            {
                return 0;
            }
            if (tungay > denngay)
            {
                sophepnam = 0;
            }
            else if (tungay == denngay)
            {
                sophepnam = phep;
            }
            else
            {
                sophepnam = TinhSoNgayPhep(tungay.Value, phep.Value, denngay);
                //if (denngay.Year > tungay.Value.Year)
                //{
                //    tungay = new DateTime(denngay.Year - 1, 12, 26);  // Đặt ngày hiện tại thành đầu năm mới
                //}
                var query = _context.ChamcongModel.Where(d => d.value == "P" && d.MANV == person.MANV);
                if (include == true)
                {

                    query = query.Where(d => d.date > tungay && d.date <= denngay);
                }
                else
                {
                    query = query.Where(d => d.date > tungay && d.date < denngay);
                }
                var count_phep = query.Include(d => d.shift).Sum(d => d.shift.factor);

                sophepnam = sophepnam - (double)count_phep;
            }
            return sophepnam;
        }
        public WorkTimeCalculator GetSchedule1(string id, TimeSpan start, TimeSpan end)
        {
            Dictionary<DayOfWeek, List<WorkShift>> MySchedule = new Dictionary<DayOfWeek, List<WorkShift>>() {
              { DayOfWeek.Sunday, new List<WorkShift>(){
                new WorkShift(){ Start = start, End = end},
              } },
              { DayOfWeek.Monday, new List<WorkShift>(){
                 new WorkShift(){ Start = start, End = end},
              } },
              { DayOfWeek.Tuesday, new List<WorkShift>(){
                 new WorkShift(){ Start = start, End = end},
              } },
              { DayOfWeek.Wednesday, new List<WorkShift>(){
                 new WorkShift(){ Start = start, End = end},
              } },
              { DayOfWeek.Thursday, new List<WorkShift>(){
                 new WorkShift(){ Start = start, End = end},
              } },
              { DayOfWeek.Friday, new List<WorkShift>(){
                 new WorkShift(){ Start = start, End = end},
              } },
              { DayOfWeek.Saturday, new List<WorkShift>(){
                 new WorkShift(){ Start = start, End = end},
              } }
            };

            var context = _context.ShiftHolidayModel.Where(d => d.shift_id == id);

            var MyHolidays = context.OrderBy(d => d.date).ToList().Select(d => new HolidayConfig() { Start = d.date.Value.Date, End = d.date.Value.Date }).ToList();

            //instantiate with configuration
            var utility = new WorkTimeCalculator(MySchedule, MyHolidays);
            return utility;
        }
        public WorkingDaysAndTimeUtility GetSchedule(string id, TimeSpan start, TimeSpan end, DateTime date_from, DateTime date_to)
        {
            var wts = new List<WorkTimeSpan>() { new WorkTimeSpan()
                { Start = start, End = end } };

            var week = new WeekDaySpan()
            {
                WorkDays = new Dictionary<DayOfWeek, WorkDaySpan>()
                {
                    {DayOfWeek.Monday, new WorkDaySpan() {TimeSpans = wts}}
                    ,
                    {DayOfWeek.Tuesday, new WorkDaySpan() {TimeSpans = wts}}
                    ,
                    {DayOfWeek.Wednesday, new WorkDaySpan() {TimeSpans = wts}}
                    ,
                    {DayOfWeek.Thursday, new WorkDaySpan() {TimeSpans = wts}}
                    ,
                    {DayOfWeek.Friday, new WorkDaySpan() {TimeSpans = wts}}
                    ,
                    {DayOfWeek.Saturday, new WorkDaySpan() {TimeSpans = wts}}
                    ,
                    {DayOfWeek.Sunday, new WorkDaySpan() {TimeSpans = wts}}
                }
            };

            var context = _context.ShiftHolidayModel.Where(d => d.shift_id == id && d.date >= date_from && d.date <= date_to);

            var holidays = context.OrderBy(d => d.date).ToList();
            //this is the configuration for holidays: 
            //in Italy we have this list of Holidays plus 1 day different on each province,
            //for mine is 1 Dec (see last element of the List<AHolyDay>).
            var italiansHoliDays = new List<AHolyDay>()
            {

            };
            italiansHoliDays.Add(new WHolidays(holidays));
            //instantiate with configuration
            var utility = new WorkingDaysAndTimeUtility(week, italiansHoliDays);
            return utility;
        }
        private static double SoLanQuaNgay26(DateTime ngayHienTai, DateTime ngayTuongLai)
        {
            double soLanQuaNgay26 = 0;

            // Duyệt qua từng tháng từ tháng hiện tại đến ngày tương lai
            DateTime current = new DateTime(ngayHienTai.Year, ngayHienTai.Month, 1);

            while (current <= ngayTuongLai)
            {
                DateTime ngay26CuaThang = new DateTime(current.Year, current.Month, 26);

                // Kiểm tra nếu ngày 26 nằm trong khoảng ngày hiện tại và ngày tương lai
                if (ngay26CuaThang > ngayHienTai && ngay26CuaThang <= ngayTuongLai)
                {
                    soLanQuaNgay26++;
                }

                // Chuyển sang tháng kế tiếp
                current = current.AddMonths(1);
            }

            return soLanQuaNgay26;
        }
        private static double TinhSoNgayPhep(DateTime ngayHienTai, double soNgayPhepHienTai, DateTime ngayTuongLai)
        {
            // Nếu ngày tương lai qua năm mới, reset ngày phép về 0
            //if (ngayTuongLai.Year > ngayHienTai.Year)
            //{
                //soNgayPhepHienTai = 1;
                //ngayHienTai = new DateTime(ngayTuongLai.Year, 1, 1);  // Đặt ngày hiện tại thành đầu năm mới
            //}

            // Tính số lần qua ngày 26 từ ngày hiện tại đến ngày tương lai
            double soLanQuaNgay26 = SoLanQuaNgay26(ngayHienTai, ngayTuongLai);

            // Tổng số ngày phép vào ngày tương lai
            return soNgayPhepHienTai + soLanQuaNgay26;
        }

    }
}