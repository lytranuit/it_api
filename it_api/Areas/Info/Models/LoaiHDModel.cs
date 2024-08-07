using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Vue.Models;

namespace Info.Models
{
    [Table("nsMALOAIHD")]
    public class LoaiHDModel
    {
        [Key]
        public string id { get; set; }
        public string MALOAIHD { get; set; }

        public string TENLOAIHD { get; set; }

    }
}
