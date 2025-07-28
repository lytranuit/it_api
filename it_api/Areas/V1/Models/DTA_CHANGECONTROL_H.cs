using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Vue.Areas.V1.Models
{
    [Table("DTA_CHANGECONTROL_H")]
    public class DTA_CHANGECONTROL_H
    {
        public string sochange { get; set; } = null!;

        public DateTime ngaydenghi { get; set; }

        public string? makhuvuc { get; set; }

        public string? tenkhuvuc { get; set; }

        public string? congthuc { get; set; }

        public string? luuhanh { get; set; }

        public string? sodangky { get; set; }

        public string? mansx { get; set; }

        public string? tennsx { get; set; }

        public string? ghichu { get; set; }

        public string? nguoidung { get; set; }
    }
}
