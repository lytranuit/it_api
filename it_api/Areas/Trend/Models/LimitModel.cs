using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Vue.Models
{
    [Table("limit")]
    public class LimitModel
    {
        [Key]
        public int id { get; set; }
        public string? name { get; set; }
        public decimal? alert_limit { get; set; }
        public decimal? action_limit { get; set; }
        public decimal? standard_limit { get; set; }
        public int? target_id { get; set; }
        public int? object_id { get; set; }
        public DateTime? date_effect { get; set; }
        public DateTime? date_from { get; set; }
        public DateTime? date_to { get; set; }
        public DateTime? deleted_at { get; set; }
        public DateTime? updated_at { get; set; }
        public DateTime? created_at { get; set; }


        [ForeignKey("object_id")]
        public virtual ObjectModel obj { get; set; }

        [ForeignKey("target_id")]
        public virtual TargetModel target { get; set; }


        public virtual List<LimitPointModel>? points { get; set; }

        public virtual List<int>? list_point
        {
            get
            {
                return points != null ? points.Select(d => d.point_id).ToList() : null;
            }
        }
    }
}
