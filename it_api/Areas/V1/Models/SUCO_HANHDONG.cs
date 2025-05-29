using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Vue.Models
{
    public class SUCO_HANHDONG
    {

        [Key]
        public int stt { get; set; }
        public string? noidung { get; set; }
        public string? trachnhiem { get; set; }
        public DateTime? dukien { get; set; }
    }
}
