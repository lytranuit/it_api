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
        public string? value { get; set; }

        public DateTime? date { get; set; }


    }
}
