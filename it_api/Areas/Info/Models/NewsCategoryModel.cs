using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Vue.Models;

namespace Info.Models
{
    [Table("a_news_category")]
    public class NewsCategoryModel
    {
        [Key]
        public int id { get; set; }
        public string news_id { get; set; }
        public string category_id { get; set; }

        [ForeignKey("news_id")]
        public NewsModel news { get; set; }

        [ForeignKey("category_id")]
        public CategoryModel category { get; set; }

    }
}
