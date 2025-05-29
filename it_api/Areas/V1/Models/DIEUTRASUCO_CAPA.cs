using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Vue.Models
{
    public class DIEUTRASUCO_CAPA
    {
        [Key]
        [Column("SOCAPA_1")]
        public string socapa { get; set; }
        [Column("NOIDUNG_")]
        public string? noidung { get; set; }
        [Column("TENBP_1")]
        public string? tenbp { get; set; }
        [Column("NGAYHOANTHANH_DUKIEN_CAPA_1")]
        public DateTime? ngayhoanthanh_dukien { get; set; }




    }
}
