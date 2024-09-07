using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Vue.Models;

namespace Info.Models
{
    [Table("a_chamcong")]
    public class ChamcongModel
    {
        [Key]
        public int id { get; set; }
        public string MANV { get; set; }
        public string NV_id { get; set; }
        public string shift_id { get; set; }

        [ForeignKey("shift_id")]
        public ShiftModel shift { get; set; }
        public string? value { get; set; }

        public string? value_new { get; set; }
        public DateTime? date { get; set; }

        public string? orderletter_id { get; set; }
        public bool? is_duyet { get; set; }

    }
}
