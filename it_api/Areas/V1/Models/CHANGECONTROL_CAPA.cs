using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Vue.Models
{
    public class CHANGECONTROL_CAPA

    {
        public string sochange { get; set; }
        public DateTime? ngaydenghi { get; set; }
        public string? anban { get; set; }

        [Key]
        public string sohanhdong { get; set; }
        public string? tenhanhdong { get; set; }
        public string? trachnhiem { get; set; }
        public DateTime? dukien { get; set; }


        public string? danhgia { get; set; }

        public string? ten { get; set; }
        public string? ten_en { get; set; }

        public int? id_muc { get; set; }
        public string? ten_muc { get; set; }
        public string? ten_en_muc { get; set; }

        public int? manhom_muc { get; set; }
        public string? tennhom_muc { get; set; }
        public string? tennhom_en_muc { get; set; }

    }
}
