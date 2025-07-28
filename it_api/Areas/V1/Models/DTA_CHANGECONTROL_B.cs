using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Vue.Areas.V1.Models
{
    [Table("DTA_CHANGECONTROL_B")]
    public class DTA_CHANGECONTROL_B
    {
        public string sochange { get; set; } = null!;

        public DateTime ngaydenghi { get; set; }

        public string? loaithaydoi { get; set; }

        public DateTime? ngaythaydoi { get; set; }

        public string? giaidoan { get; set; }

        public string? maphong { get; set; }

        public string? tenphong { get; set; }

        public string? capsach { get; set; }

        public string? thietbido { get; set; }

        public string? thaydoinhaxuong { get; set; }

        public string? anhhuongtienich { get; set; }

        public string? phanmemdikem { get; set; }

        public string? tacdongsanpham { get; set; }

        public string? ghichu { get; set; }

        public string? nguoidung { get; set; }
    }
}
