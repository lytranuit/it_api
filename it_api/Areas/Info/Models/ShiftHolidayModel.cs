using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Vue.Models;

namespace Info.Models
{
    [Table("a_shift_holiday")]
    public class ShiftHolidayModel
    {
        [Key]
        public string id { get; set; }
        public string? name { get; set; }

        public DateTime? date { get; set; }

        public string shift_id { get; set; }

    }
}
