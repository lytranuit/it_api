using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Vue.Models
{
    public class DIEUTRASUCO_PARTB
    {
        [Key]
        public int id { get; set; }

        public string? hoten { get; set; }
        public string? tenbp { get; set; }
        public string? nguyennhancotloi { get; set; }



    }
}
