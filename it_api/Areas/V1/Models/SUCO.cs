using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Vue.Models
{
    public class SUCO
    {
        [Key]
        public int id { get; set; }
        public string so { get; set; }
        public string? tieude { get; set; }
        public bool? suco { get; set; }


        public string? suco_baolau { get; set; }
        public string? suco_khinao { get; set; }
        public string? suco_phathien { get; set; }
        public string? suco_odau { get; set; }
        public string? suco_mabp { get; set; }


        public string? suco_anhhuong { get; set; }
        public string? suco_anhhuong_sanpham { get; set; }
        public string? sanpham { get; set; }
        public string? giaitrinhtre { get; set; }

        public string? mabp_02 { get; set; }
        public string? mabp_05 { get; set; }
        public string? mabp_06 { get; set; }
        public string? mabp_07 { get; set; }
        public string? mabp_09 { get; set; }
        public string? mabp_10 { get; set; }
        public string? mabp_11 { get; set; }
        public string? mabp_13 { get; set; }
        public string? mabp_16 { get; set; }


        public string? tenkhuvuc { get; set; }
        public string? tenkhuvuc_VN { get; set; }



    }
}
