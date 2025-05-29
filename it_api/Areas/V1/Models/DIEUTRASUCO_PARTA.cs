using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Vue.Models
{
    public class DIEUTRASUCO_PARTA
    {
        [Key]
        [Column("id_id")]
        public int id { get; set; }

        public string? hanhdong { get; set; }

        public string? tenbp { get; set; }
        public string? nguyennhancotloi { get; set; }
        public DateTime? ngayhoanthanh_dukien { get; set; }

    }
}
