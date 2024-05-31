using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Holdtime.Models
{
    [Table("stage")]
    public class StageModel
    {
        [Key]
        public int id { get; set; }
        public string? name { get; set; }
        public DateTime? deleted_at { get; set; }
        public DateTime? updated_at { get; set; }
        public DateTime? created_at { get; set; }
    }
}
