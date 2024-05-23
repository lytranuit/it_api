using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Vue.Models
{
    [Table("object_target")]
    public class ObjectTargetModel
    {
        [Key]
        public int id { get; set; }
        public int object_id { get; set; }
        public int target_id { get; set; }
        public int order { get; set; }

        [ForeignKey("object_id")]
        public virtual ObjectModel obj { get; set; }
        [ForeignKey("target_id")]
        public virtual TargetModel target { get; set; }
    }
}
