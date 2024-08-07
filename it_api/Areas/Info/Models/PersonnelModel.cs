using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Vue.Models;

namespace Info.Models
{
    [Table("nsHOSONHANVIEN")]
    public class PersonnelModel
    {
        [Key]
        public string id { get; set; }
        public string MANV { get; set; }
        public string? HOVATEN { get; set; }
        public string? GIOITINH { get; set; }
        public string? DANTOC { get; set; }
        public string? QUOCTICH { get; set; }
        public string? EMAIL { get; set; }
        public string? DIENTHOAI { get; set; }
        public DateTime? NGAYSINH { get; set; }
        public string? NOISINH { get; set; }
        public string? NGUYENQUAN { get; set; }
        public string? THUONGTRU { get; set; }
        public string? THUONGTRU_EN { get; set; }
        public DateTime? NGAYNHANVIEC { get; set; }
        public DateTime? NGAYHOCVIEC { get; set; }
        public DateTime? NGAYTHUVIEC { get; set; }
        public DateTime? NGAYKYHD { get; set; }
        public DateTime? NGAYKTHD { get; set; }
        public string? SOCCCD { get; set; }
        public DateTime? NGAYCAPCCCD { get; set; }
        public string? NOICAPCCCD { get; set; }
        public string? SOBHXH { get; set; }
        public DateTime? NGAYCAPBHXH { get; set; }
        public string? NOICAPBHXH { get; set; }
        public string? SOHD { get; set; }
        public string? tinhtrang { get; set; }
        public DateTime? NGAYNGHITHAISAN { get; set; }
        public DateTime? NGAYNGHIVIEC { get; set; }
        public string? lydonghiviec { get; set; }
        public string? sotk_icb { get; set; }
        public string? sotk_vba { get; set; }
        public string? MATHUE { get; set; }
        public string? MATRINHDO { get; set; }
        public string? CHUYENMON { get; set; }
        public string? MAPHONG { get; set; }
        public string? MAKHUVUC { get; set; }
        public string? MACHUCVU { get; set; }
        public string? LOAIHD { get; set; }

        public string? NGUOIPHANCONG { get; set; }
        public string? DIADIEM { get; set; }
        public string? CONGVIEC { get; set; }
    }
}
