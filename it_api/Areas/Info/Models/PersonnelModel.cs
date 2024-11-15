using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Text;
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
        public string? MAQUANLYTRUCTIEP { get; set; }
        public string? DIADIEM { get; set; }
        public string? CONGVIEC { get; set; }
        public string? MACC { get; set; }

        public bool? autoeat { get; set; }
        public double? ngayphep { get; set; }
        public DateTime? ngayphep_date { get; set; }

        public double? tien_luong { get; set; }

        public double? tien_luong_dot1 { get; set; }

        public double? tien_luong_kpi { get; set; }

        public double? tong_thunhap { get; set; }
        public int? nguoiphuthuoc { get; set; }

        public double? tyle_bhxh { get; set; }

        public double? tyle_bhyt { get; set; }

        public double? tyle_dpcd { get; set; }

        public double? tyle_bhtn { get; set; }

        public double? cty_bhxh { get; set; }

        public double? cty_bhyt { get; set; }

        public double? cty_dpcd { get; set; }

        public double? cty_bhtn { get; set; }


        public double? pc_thamnien { get; set; }
        public double? pc_khuvuc { get; set; }
        public double? pc_thuhut { get; set; }
        public double? pc_thuebang { get; set; }
        public double? pc_trachnhiem { get; set; }
        public double? pc_khac { get; set; }
        public double? pc { get; set; }
        public double? pc_hieusuat { get; set; }

        public bool? is_bhxh { get; set; }
        public bool? is_thue { get; set; }
        public string? image_url { get; set; }
        [NotMapped]
        public List<string>? list_shift { get; set; }

        public string NormalizeName
        {
            get
            {
                // Chuyển về chữ thường
                string normalized = HOVATEN.ToLower();

                // Loại bỏ dấu tiếng Việt
                if (string.IsNullOrWhiteSpace(normalized))
                    return normalized;
                normalized = normalized.Replace('đ', 'd').Replace('Đ', 'D');
                normalized = normalized.Normalize(NormalizationForm.FormD);
                var chars = normalized.ToCharArray();
                StringBuilder sb = new StringBuilder();

                foreach (var c in chars)
                {
                    UnicodeCategory unicodeCategory = CharUnicodeInfo.GetUnicodeCategory(c);
                    if (unicodeCategory != UnicodeCategory.NonSpacingMark)
                    {
                        sb.Append(c);
                    }
                }
                normalized = sb.ToString().Normalize(NormalizationForm.FormC);
                // Xóa khoảng trắng thừa
                normalized = Regex.Replace(normalized, @"\s+", " ").Trim();

                return normalized;
            }
        }
    }

}
