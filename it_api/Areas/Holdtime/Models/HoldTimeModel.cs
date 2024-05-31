using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Holdtime.Models
{
    [Table("hold_time")]
    public class HoldTimeModel
    {
        [Key]
        public int id { get; set; }
        public int? hold_id { get; set; }
        public string? name { get; set; }
        public string? type { get; set; }
        public int? time { get; set; }
        public string? time_type { get; set; }
        public string? based { get; set; }

        public DateTime? date_theory { get; set; }
        public DateTime? date_reality { get; set; }

        public string? note { get; set; }
        public int? num_get { get; set; }
        public bool? is_pass { get; set; }
        public virtual List<HoldTimeTargetModel>? targets { get; set; }


        [ForeignKey("hold_id")]
        public virtual HoldModel Hold { get; set; }
        [NotMapped]
        public List<int>? list_target { get; set; }

        public DateTime? deleted_at { get; set; }
        public DateTime? updated_at { get; set; }
        public DateTime? created_at { get; set; }
    }
}
