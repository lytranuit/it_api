using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Vue.Models
{
    public class COAModel
    {
        [Key]
        public int id { get; set; }
        public int? id_parent { get; set; }

        public string solenh { get; set; }
        public DateTime ngaylenh { get; set; }

        public string? chitieu { get; set; }
        public string? chitieu_en { get; set; }
        public string? tieuchuan { get; set; }
        public string? tieuchuan_en { get; set; }
        public double? thucte { get; set; }
        public string? thucte_text { get; set; }
        public string? thucte_text_en { get; set; }
        public string? mota { get; set; }
        public string? mota_en { get; set; }
        public string? type { get; set; }




        public string? mahh { get; set; }
        public string? tenhh { get; set; }
        public string? tenhoatchat { get; set; }
        public string? dangbaoche { get; set; }
        public DateTime? ngaysx { get; set; }
        public string? quicachdonggoi { get; set; }
        public string? sophieu { get; set; }

        public string? malo { get; set; }
        public string? sodk { get; set; }


        public string? handung { get; set; }
        public string? theotieuchuan { get; set; }
        public string? ghichu { get; set; }
        public string? ketluan { get; set; }
        public string? sop { get; set; }
        public string? ngaysop { get; set; }

    }
}
