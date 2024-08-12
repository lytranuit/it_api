using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Vue.Models;

namespace Info.Models
{
    [Table("a_calendar_holiday")]
    public class CalendarHolidayModel
    {
        [Key]
        public string id { get; set; }
        public string? name { get; set; }

        public DateTime? date { get; set; }

        public string calendar_id { get; set; }

    }
}
