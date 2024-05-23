using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Vue.Models
{
    [Table("result")]
    public class ResultModel
    {
        [Key]
        public int id { get; set; }
        public int? point_id { get; set; }
        public decimal? value { get; set; }
        public string? value_text { get; set; }
        public string? note { get; set; }
        public DateTime? date { get; set; }
        public int? object_id { get; set; }
        public int? target_id { get; set; }
        public DateTime? deleted_at { get; set; }
        public DateTime? updated_at { get; set; }
        public DateTime? created_at { get; set; }

        public decimal? limit_action { get; set; }
        public decimal? limit_alert { get; set; }


        public int? limit_id { get; set; }

        //public virtual List<LimitPointModel> LimitPoints { get; set; }

        [ForeignKey("object_id")]
        public virtual ObjectModel obj { get; set; }

        [ForeignKey("target_id")]
        public virtual TargetModel target { get; set; }

        [ForeignKey("point_id")]
        public virtual PointModel point { get; set; }
        [ForeignKey("limit_id")]
        public virtual LimitModel LimitModel { get; set; }

    }
}
