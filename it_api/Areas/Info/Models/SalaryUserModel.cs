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
        public string? MABOPHAN { get; set; }
        public string? tinhtrangNV { get; set; }

        public decimal? ngaycongchuan { get; set; }
        public decimal? ngaycongthucte { get; set; }
        public decimal? luongcb { get; set; }
        public decimal? luongdongbhxh { get; set; }
        public decimal? luongdoanhso { get; set; }


        public decimal? tc_hieusuat { get; set; }
        public decimal? tc_thuebang { get; set; }
        public decimal? tc_xangxe { get; set; }
        public decimal? tc_thamnien { get; set; }
        public decimal? tc_thuhut { get; set; }
        public decimal? tc_khuvuc { get; set; }
        public decimal? tc_chucvu { get; set; }
        public decimal? tc_tienan { get; set; }
        public decimal? tc_khac { get; set; }
        public decimal? tong_tc { get; set; }


        public decimal? luongkpi { get; set; }
        public decimal? bosung { get; set; }
        public decimal? tongthunhap { get; set; }
        public decimal? tong_tn { get; set; }



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
        public decimal? tongphep { get; set; }
        public decimal? phepconlai { get; set; }
        public decimal? phepdauky { get; set; }
        public decimal? khoantru { get; set; }
        public decimal? khoancong { get; set; }
        public decimal? khoantru_sauthue { get; set; }
        public decimal? khoancong_sauthue { get; set; }

        public string? note_khoantru { get; set; }
        public string? note_khoancong { get; set; }
        public string? note_khoantru_sauthue { get; set; }
        public string? note_khoancong_sauthue { get; set; }
        public string? note { get; set; }

        public decimal? tyle_bhxh { get; set; }
        public decimal? tyle_bhyt { get; set; }
        public decimal? tyle_bhtn { get; set; }
        public decimal? tyle_dpcd { get; set; }


        public decimal? tyle_bhxh_cty { get; set; }
        public decimal? tyle_bhyt_cty { get; set; }
        public decimal? tyle_bhtn_cty { get; set; }
        public decimal? tyle_dpcd_cty { get; set; }

        public bool? is_bhxh { get; set; }

        public bool? is_thue { get; set; }

        public Xeploai? xeploai { get; set; } = Xeploai.None;
        public decimal? luongxeploai { get; set; }

    }

    public enum Xeploai
    {
        None = 0,
        A = 1,
        B = 2,
        C = 3,
        D = 4,
        E = 5,
    }
}
