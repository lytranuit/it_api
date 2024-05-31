using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Vue.Models;

namespace Holdtime.Models
{
    [Table("hold_alert")]
    public class HoldAlertModel
    {
        [Key]
        public int id { get; set; }
        public int? hold_id { get; set; }
        public string? user_id { get; set; }

        [ForeignKey("hold_id")]
        public virtual HoldModel Hold { get; set; }

        [ForeignKey("user_id")]
        public virtual UserModel user { get; set; }
    }
}
