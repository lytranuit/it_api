using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Vue.Models;

namespace Info.Models
{
    [Table("a_salary_user")]
    public class SalaryUserModel
    {
        [Key]
        public int id { get; set; }
        public string? email { get; set; }
        public string? person_id { get; set; }

        public string? file_url { get; set; }
        public string? salary_id { get; set; }


        [ForeignKey("salary_id")]
        public SalaryModel salary { get; set; }




        public string? MANV { get; set; }
        public string? HOVATEN { get; set; }
        public string? CHUCVU { get; set; }
        public string? BOPHAN { get; set; }

        public decimal? ngaycongchuan { get; set; }
        public decimal? ngaycongthucte { get; set; }
        public decimal? luongcb { get; set; }
        public decimal? luongdongbhxh { get; set; }


        public decimal? tc_trangphuc { get; set; }
        public decimal? tc_dienthoai { get; set; }
        public decimal? tc_xangxe { get; set; }
        public decimal? tc_antrua { get; set; }
        public decimal? tc_chuyencan { get; set; }
        public decimal? tc_anca { get; set; }
        public decimal? tc_chucvu { get; set; }
        public decimal? tong_tc { get; set; }


        public decimal? luongkpi { get; set; }
        public decimal? bosung { get; set; }
        public decimal? tongthunhap { get; set; }


        public decimal? thunhapchiuthue { get; set; }
        public decimal? tncn_banthan { get; set; }
        public int? tncn_songuoiphuthuoc { get; set; }
        public decimal? tncn_nguoiphuthuoc { get; set; }
        public decimal? tncn_bhxh { get; set; }
        public decimal? thunhaptinhthue { get; set; }
        public decimal? thue_tncn { get; set; }
        public decimal? dpcd { get; set; }
        public decimal? thuclanh { get; set; }
        public decimal? tamungdot1 { get; set; }
        public decimal? conlai { get; set; }

    }
}
