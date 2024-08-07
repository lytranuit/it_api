using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Vue.Models;

namespace Info.Models
{
    [Table("a_hot_news")]
    public class HotNewsModel
    {
        [Key]
        public int id { get; set; }
        public string message { get; set; }
     
    }
}
