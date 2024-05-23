using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Vue.Models
{
    [Table("object")]
    public class ObjectModel
    {
        [Key]
        public int id { get; set; }
        public string? name { get; set; }
        public string? name_en { get; set; }
        public bool? is_multi_target { get; set; }
        public DateTime? deleted_at { get; set; }
        public DateTime? updated_at { get; set; }
        public DateTime? created_at { get; set; }

        public virtual List<ObjectTargetModel> targets { get; set; }
    }
}
