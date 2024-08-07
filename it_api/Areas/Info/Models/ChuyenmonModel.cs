using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Vue.Models;

namespace Info.Models
{
    [Table("nsMACHUYENMON")]
    public class ChuyenmonModel
    {
        [Key]
        public string id { get; set; }
        public string? MACHUYENMON { get; set; }

        public string? TENCHUYENMON { get; set; }

    }
}
