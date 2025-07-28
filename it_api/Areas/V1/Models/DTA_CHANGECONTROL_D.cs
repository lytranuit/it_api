using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Vue.Areas.V1.Models
{
    [Table("DTA_CHANGECONTROL_D")]
    public class DTA_CHANGECONTROL_D
    {
        public string sochange { get; set; } = null!;

        public DateTime ngaydenghi { get; set; }

        public string? loaithaydoi { get; set; }

        public string? mucdich { get; set; }

        public string? ghichu { get; set; }

        public string? nguoidung { get; set; }
    }
}
