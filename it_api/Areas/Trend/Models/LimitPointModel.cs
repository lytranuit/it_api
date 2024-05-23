using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Vue.Models
{
    [Table("limit_point")]
    public class LimitPointModel
    {
        [Key]
        public int id { get; set; }
        public int limit_id { get; set; }
        public int point_id { get; set; }
        [ForeignKey("limit_id")]
        public virtual LimitModel limit { get; set; }
        [ForeignKey("point_id")]
        public virtual PointModel point { get; set; }
    }
}
