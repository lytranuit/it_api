using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Vue.Models;

namespace Info.Models
{
    [Table("nsMAPHONG")]
    public class PhongModel
    {
        [Key]
        public string id { get; set; }
        public string? MAPHONG { get; set; }

        public string? TENPHONG { get; set; }
        public string? MAKHUVUC { get; set; }
        public string? truongbophan_id { get; set; }
        public string? quanlycong_id { get; set; }
        


    }
}
