using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Vue.Models;

namespace Info.Models
{
    [Table("a_calendar")]
    public class CalendarModel
    {
        [Key]
        public string id { get; set; }
        public string name { get; set; }

        public DateTime? date { get; set; }


        public DateTime? deleted_at { get; set; }
        public DateTime? updated_at { get; set; }
        public DateTime? created_at { get; set; }

    }
}
