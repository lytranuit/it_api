using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Vue.Models;

namespace Info.Models
{
    [Table("a_phonghop")]
    public class PhonghopModel
    {
        [Key]
        public int id { get; set; }
        public string name { get; set; }



    }
}
