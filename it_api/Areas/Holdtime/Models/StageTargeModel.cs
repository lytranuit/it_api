using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Holdtime.Models
{
    [Table("stage_target")]
    public class StageTargetModel
    {
        [Key]
        public int id { get; set; }
        public int? stage_id { get; set; }
        public int? target_id { get; set; }
    }
}
