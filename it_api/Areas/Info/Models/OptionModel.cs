using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Vue.Models;

namespace Info.Models
{
    [Table("a_options")]
    public class OptionModel
    {
        [Key]
        public int id { get; set; }
        public string key { get; set; }

        public DateTime? date_value { get; set; }
        public string? value { get; set; }


    }
}
