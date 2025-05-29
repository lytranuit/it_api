using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Vue.Models
{
    public class CHANGECONTROL_ANBAN_IN

    {
        //[Key]
        public string makhuvuc { get; set; }
        public DateTime ngayhieuluc { get; set; }
        public string anban { get; set; }


        public string mota { get; set; }
        public string? mota_en { get; set; }


    }
}
