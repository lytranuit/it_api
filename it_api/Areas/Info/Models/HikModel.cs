using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Vue.Models;

namespace Info.Models
{
    [Table("a_hik")]
    public class HikModel
    {
        [Key]
        public string id { get; set; }
        public string person_name { get; set; }
        public string device { get; set; }
        public string deviceno { get; set; }
        public string card { get; set; }

        public DateTime? date { get; set; }
        public TimeSpan? time { get; set; }
        public DateTime? datetime { get; set; }



    }
}
