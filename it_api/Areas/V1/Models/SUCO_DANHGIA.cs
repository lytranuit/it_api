using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Vue.Models
{
    public class SUCO_DANHGIA
    {
        [Key]
        [Column("mabp_1")]
        public string mabp { get; set; }
        public string? tenkhuvuc_1 { get; set; }
        public string? tenkhuvuc_VN_1 { get; set; }
        public string? hanhdong { get; set; }
        public string? dieutra { get; set; }
        public bool? tacdong { get; set; }
        public string? giaithich { get; set; }
        public string? tacdongchitiet { get; set; }
        public string? dinhkem { get; set; }

    }
}
