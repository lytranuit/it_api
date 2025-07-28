using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Vue.Areas.V1.Models
{
    [Table("DTA_CHANGECONTROL_G")]
    public class DTA_CHANGECONTROL_G
    {
        public string sochange { get; set; } = null!;

        public DateTime ngaydenghi { get; set; }

        public string? loaithaydoi { get; set; }

        public string? mansx { get; set; }

        public string? tennsx { get; set; }

        public string? mucdich { get; set; }

        public string? danhgia { get; set; }

        public string? dulieudanhgia { get; set; }

        public string? dinhluong { get; set; }

        public string? khutrung { get; set; }

        public string? sudung { get; set; }

        public string? ghichu { get; set; }

        public string? nguoidung { get; set; }
    }
}
