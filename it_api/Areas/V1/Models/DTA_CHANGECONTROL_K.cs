using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Vue.Areas.V1.Models
{
    [Table("DTA_CHANGECONTROL_K")]
    public class DTA_CHANGECONTROL_K
    {
        public string sochange { get; set; } = null!;

        public DateTime ngaydenghi { get; set; }

        public string? loaithaydoi { get; set; }

        public string? thaydoitieuchuan { get; set; }

        public string? nghiencuudoondinh { get; set; }

        public string? ghichu { get; set; }

        public string? nguoidung { get; set; }
    }
}
