using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Vue.Models;

namespace Info.Models
{
    [Table("a_shift")]
    public class ShiftModel
    {
        [Key]
        public string id { get; set; }
        public string name { get; set; }
        public string code { get; set; }

        public DateTime? date { get; set; }

        public TimeSpan? time_from { get; set; }
        public TimeSpan? time_to { get; set; }
        public decimal? factor { get; set; }


        public DateTime? deleted_at { get; set; }
        public DateTime? updated_at { get; set; }
        public DateTime? created_at { get; set; }

    }
}
