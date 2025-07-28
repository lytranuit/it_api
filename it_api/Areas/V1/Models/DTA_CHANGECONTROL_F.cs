using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Vue.Areas.V1.Models
{
    [Table("DTA_CHANGECONTROL_F")]
    public class DTA_CHANGECONTROL_F
    {
        public string sochange { get; set; } = null!;

        public DateTime ngaydenghi { get; set; }

        public string? loaithaydoi { get; set; }

        public string? thaydoisanpham { get; set; }

        public string? dulieudoondinh { get; set; }

        public string? ghichu { get; set; }

        public string? nguoidung { get; set; }
    }
}
