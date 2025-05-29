using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Vue.Models
{
    public class CHANGECONTROL_PHANLOAI
    {
        [Key]
        public string sochange { get; set; }
        public DateTime? ngaydenghi { get; set; }
        public string? anban { get; set; }
        public DateTime? ngayhieuluc { get; set; }


        public string? anhhuongchatluong_thap { get; set; }
        public string? anhhuongchatluong_trungbinh { get; set; }
        public string? anhhuongchatluong_cao { get; set; }


        public string? anhhuonghoso_thap { get; set; }
        public string? anhhuonghoso_trungbinh { get; set; }
        public string? anhhuonghoso_cao { get; set; }


        public string? thaydoi_thap { get; set; }
        public string? thaydoi_trungbinh { get; set; }
        public string? thaydoi_cao { get; set; }
        public string? phanloai { get; set; }

    }
}
