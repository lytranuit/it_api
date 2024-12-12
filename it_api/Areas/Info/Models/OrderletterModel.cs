using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Vue.Models;

namespace Info.Models
{
    [Table("a_order_letter")]
    public class OrderletterModel
    {
        [Key]
        public string id { get; set; }
        public string name { get; set; }
        public string description { get; set; }
        public int? type { get; set; }
        public string? created_by { get; set; }

        [ForeignKey("created_by")]
        public UserModel? user_created_by { get; set; }
        public int? status_id { get; set; }
        public int? status1_id { get; set; }

        public string? user_accept_id { get; set; }
        [ForeignKey("user_accept_id")]
        public UserModel? user_accept { get; set; }
        public string? user1_accept_id { get; set; }
        [ForeignKey("user1_accept_id")]
        public UserModel? user1_accept { get; set; }

        public string? user_id { get; set; }

        [ForeignKey("user_id")]
        public UserModel? user { get; set; }

        public string? note { get; set; }
        public string? note1 { get; set; }

        public DateTime? date_accept { get; set; }
        public DateTime? date1_accept { get; set; }
        public DateTime? deleted_at { get; set; }
        public DateTime? updated_at { get; set; }
        public DateTime? created_at { get; set; }

    }

    enum OrderletterStatus
    {
        [Display(Name = "Tạo mới")]
        New = 1,
        [Display(Name = "Duyệt")]
        Duyet = 2,
        [Display(Name = "Không duyệt")]
        Khong_duyet = 3,

    }
}
