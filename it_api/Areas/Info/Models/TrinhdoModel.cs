using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Vue.Models;

namespace Info.Models
{
    [Table("nsMATRINHDO")]
    public class TrinhdoModel
    {
        [Key]
        public string id { get; set; }
        public string MATRINHDO { get; set; }

        public string TENTRINHDO { get; set; }

    }
}
