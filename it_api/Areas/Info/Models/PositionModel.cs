using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Vue.Models;

namespace Info.Models
{
    [Table("nsMACHUCVU")]
    public class PositionModel
    {
        [Key]
        public string id { get; set; }
        public string? MACHUCVU { get; set; }

        public string TENCHUCVU { get; set; }
        public int? sort { get; set; }

    }
}
