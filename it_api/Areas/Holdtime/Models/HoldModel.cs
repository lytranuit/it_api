using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Holdtime.Models
{
    [Table("hold")]
    public class HoldModel
    {
        [Key]
        public int id { get; set; }
        public string? tensp { get; set; }
        public string? masp { get; set; }
        public string? malo { get; set; }
        public string? manghiencuu { get; set; }
        public string? maphantich { get; set; }
        public string? madecuong { get; set; }
        public DateTime? date_start { get; set; }
        public DateTime? date_manufacture { get; set; }
        public int? amount { get; set; }
        public int? remain { get; set; }
        public int? stage_id { get; set; }

        public virtual List<HoldTimeModel>? times { get; set; }
        [ForeignKey("stage_id")]
        public virtual StageModel stage { get; set; }
        public DateTime? deleted_at { get; set; }
        public DateTime? updated_at { get; set; }
        public DateTime? created_at { get; set; }
    }
}
