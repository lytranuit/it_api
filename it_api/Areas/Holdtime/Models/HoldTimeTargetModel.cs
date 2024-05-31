using System;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Vue.Models;

namespace Holdtime.Models
{
    [Table("hold_time_target")]
    public class HoldTimeTargetModel
    {
        [Key]
        public int id { get; set; }
        public int? hold_time_id { get; set; }
        public int? target_id { get; set; }
        public bool? is_pass { get; set; }

        [ForeignKey("hold_time_id")]
        public virtual HoldTimeModel HoldTime { get; set; }

        [ForeignKey("target_id")]
        public virtual Holdtime.Models.TargetModel target { get; set; }
    }
}
