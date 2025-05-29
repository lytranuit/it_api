using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Vue.Models;

namespace Info.Models
{
    [Table("TBL_DANHMUCNGUOIPHUTHUOC")]
    public class NguoiphuthuocModel
    {
        [Key]
        public int id { get; set; }
        public string MANV { get; set; }

        public string? tennguoiphuthuoc { get; set; }
        public string? namsinh { get; set; }
        public string? quanhe { get; set; }
        public string? masothue { get; set; }
        [NotMapped]
        public DateTime? ngaysinh
        {
            get
            {
                if (namsinh != null)
                {
                    try
                    {
                        DateTime myDate = DateTime.ParseExact(namsinh, "dd/MM/yyyy",
                                           System.Globalization.CultureInfo.InvariantCulture);
                        return myDate;

                    }
                    catch (Exception ex)
                    {
                        return null;

                    }
                }
                else
                {
                    return null;
                }

            }
            set
            {
                if (value != null && value.Value.Kind == DateTimeKind.Utc)
                {
                    value = value.Value.ToLocalTime();
                }
                namsinh = value != null ? value.Value.ToString("dd/MM/yyyy") : null;
            }
        }
        public bool? is_phuthuoc { get; set; }

        // Navigation property
        public PersonnelModel? PersonnelModel { get; set; }

    }
}
