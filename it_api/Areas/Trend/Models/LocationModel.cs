using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Vue.Models
{
    [Table("location")]
    public class LocationModel
    {
        [Key]
        public int id { get; set; }
        public string? code { get; set; }
        public string? name { get; set; }
        public string? name_en { get; set; }
        public string? color { get; set; }
        public int? parent { get; set; }
        public int? stt { get; set; }
        public int? count_child { get; set; }
        public bool? is_expand { get; set; }
        public DateTime? deleted_at { get; set; }
        public DateTime? updated_at { get; set; }
        public DateTime? created_at { get; set; }
        [NotMapped]
        public List<LocationModel>? children { get; set; }
    }
}
