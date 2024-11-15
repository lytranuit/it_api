using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Vue.Models
{
    public class LenhXuatXuongModel
    {
        [Key]

        public string solenh { get; set; }
        public DateTime ngaylenh { get; set; }

        public string? mapl { get; set; }
        public string? tenpl { get; set; }




        public string? mahh { get; set; }
        public string? tenhh { get; set; }
        public string? mahh_goc { get; set; }
        public string? tenhh_goc { get; set; }
        public string? colo { get; set; }
        public string? tenhoatchat { get; set; }
        public string? dangbaoche { get; set; }
        public DateTime? ngaysx { get; set; }
        public string? quicachdonggoi { get; set; }
        public string? dvt { get; set; }
        //public string? sophieu { get; set; }
        public decimal? soluongdonggoi { get; set; }

        public string? malo_goc { get; set; }
        public string? malo { get; set; }
        public string? sodk { get; set; }
        public string? coa { get; set; }


        public string? handung { get; set; }
        public string? tieuchuan { get; set; }

        public string? donvi { get; set; }
        public string? donvi_en { get; set; }
        public string? donvi_thung { get; set; }
        public string? donvi_thung_en { get; set; }
        public string? donvi_thungle { get; set; }
        public string? donvi_thungle_en { get; set; }

        public string? sop { get; set; }
        public string? ngaysop { get; set; }

        public decimal? thung1 { get; set; }
        public decimal? thung2 { get; set; }
        public decimal? hop1 { get; set; }
        public decimal? hop2 { get; set; }
        public decimal? tong1
        {
            get
            {
                return hop1 + hop2;
            }
        }
        public decimal? tong2 { get; set; }
        public decimal? vi2 { get; set; }
    }
}
