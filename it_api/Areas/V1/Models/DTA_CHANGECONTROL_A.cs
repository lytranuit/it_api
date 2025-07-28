using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Vue.Areas.V1.Models
{
    [Table("DTA_CHANGECONTROL_A")]
    public class DTA_CHANGECONTROL_A
    {
        public string sochange { get; set; } = null!;

        public DateTime ngaydenghi { get; set; }

        public string? loaithaydoi { get; set; }

        public DateTime? ngaythaydoi { get; set; }

        public string? maphong { get; set; }

        public string? tenphong { get; set; }

        public string? mucdich { get; set; }

        public string? thaydoibanve { get; set; }

        public string? thaydoiduongdi { get; set; }

        public string? capsach { get; set; }

        public string? thuchien { get; set; }

        public string? thaydoitienich { get; set; }

        public string? ghichu { get; set; }

        public string? nguoidung { get; set; }
    }
}
