using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Vue.Models;

namespace Info.Models
{
    [Table("a_news")]
    public class NewsModel
    {
        [Key]
        public int id { get; set; }
        public string title { get; set; }
        public string content { get; set; }

        public bool? is_publish { get; set; }
        public string? created_by { get; set; }


        [ForeignKey("created_by")]
        public virtual UserModel? user_created_by { get; set; }

        public DateTime? deleted_at { get; set; }
        public DateTime? updated_at { get; set; }
        public DateTime? created_at { get; set; }
    }
}
