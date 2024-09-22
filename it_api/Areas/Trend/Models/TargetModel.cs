using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Vue.Models
{
    [Table("Target")]
    public class TargetModel
    {
        [Key]
        public int id { get; set; }
        public string? name { get; set; }
        public string? name_en { get; set; }
        public string? unit { get; set; }
        public string? value_type { get; set; }
		public string? type { get; set; }
		public string? text_yes { get; set; }
        public string? text_no { get; set; }
        public string? text_yes_en { get; set; }
        public string? text_no_en { get; set; }
        public DateTime? deleted_at { get; set; }
        public DateTime? updated_at { get; set; }
        public DateTime? created_at { get; set; }
    }
}
