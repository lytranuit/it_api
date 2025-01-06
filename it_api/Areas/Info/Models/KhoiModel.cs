using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Vue.Models;

namespace Info.Models
{
    [Table("nsMAKHUVUC")]
    public class KhoiModel
    {
        [Key]
        public string id { get; set; }
        public string MAKHUVUC { get; set; }

        public string TENKHUVUC { get; set; }
        public int? sort { get; set; }

    }
}
