using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Vue.Models;

namespace Info.Models
{
    [Table("nsNGANHANG")]
    public class NganhangModel
    {
        [Key]
        public string id { get; set; }
        public string nganhang { get; set; }

    }
}
