using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Vue.Models;

namespace Info.Models
{
    [Table("a_shift_user")]
    public class ShiftUserModel
    {
        [Key]
        public int id { get; set; }
        public string? person_id { get; set; }
        public string? email { get; set; }

        public string? shift_id { get; set; }


        [ForeignKey("shift_id")]
        public ShiftModel shift { get; set; }

    }
}
