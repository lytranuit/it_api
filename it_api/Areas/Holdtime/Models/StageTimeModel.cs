using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Holdtime.Models
{
    [Table("stage_time")]
    public class StageTimeModel
    {
        [Key]
        public int id { get; set; }
        public int? stage_id { get; set; }
        public int? time { get; set; }
        public string? time_type { get; set; }
        public string? based { get; set; }

    }
}
