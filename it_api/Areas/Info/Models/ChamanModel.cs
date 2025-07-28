using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Vue.Models;

namespace Info.Models
{
    [Table("a_chaman")]
    public class ChamanModel
    {
        [Key]
        public string id { get; set; }
        public string MANV { get; set; }
        public string NV_id { get; set; }
        public bool? value { get; set; }

        public DateTime? date { get; set; }
        public string calendar_id { get; set; }

        [NotMapped]
        public HikModel first_hik { get; set; }
        [NotMapped]
        public int trangthai { get; set; }

    }
}
