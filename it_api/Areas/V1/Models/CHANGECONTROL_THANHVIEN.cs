using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Vue.Models
{
    public class CHANGECONTROL_THANHVIEN
    {
        public string sochange { get; set; }
        public DateTime? ngaydenghi { get; set; }
        public string? anban { get; set; }

        public string? mabp { get; set; }
        [Key]
        public string hoten { get; set; }
        public string? tenbp { get; set; }


    }
}
