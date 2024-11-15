using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Vue.Models;

namespace Info.Models
{
    [Table("nsMATINH")]
    public class TinhModel
    {
        [Key]
        public string id { get; set; }
        public string MaTinh { get; set; }

        public string? TenTinh { get; set; }

    }
}
