using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Vue.Models;

namespace Info.Models
{
    [Table("a_category")]
    public class CategoryModel
    {
        [Key]
        public string id { get; set; }
        public string name { get; set; }

        public string? created_by { get; set; }


        [ForeignKey("created_by")]
        public virtual UserModel? user_created_by { get; set; }

        //public virtual List<NewsCategoryModel> list_news { get; set; }

        public DateTime? deleted_at { get; set; }
        public DateTime? updated_at { get; set; }
        public DateTime? created_at { get; set; }


        [NotMapped]
        public virtual List<NewsModel> list_news { get; set; }
    }
}
